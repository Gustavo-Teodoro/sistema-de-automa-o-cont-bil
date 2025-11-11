"""
Módulo de automação Selenium para o Sistema de Contabilidade.
Responsável por automatizar o preenchimento de dados no site.
"""

import time
import logging
from pathlib import Path
from typing import Callable, Optional, List, Dict, Any
from dataclasses import dataclass

import openpyxl
import config
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


logger = logging.getLogger(__name__)


@dataclass
class Product:
    """Classe para representar um produto."""
    cliente: str
    produto: str
    quantidade: int
    categoria: str


class SeleniumAutomation:
    """Classe principal para automação de preenchimento de formulários."""
    
    def __init__(
        self,
        url: str = "https://devaprender-contabil.netlify.app/",
    headless: bool = config.HEADLESS_MODE,
        progress_callback: Optional[Callable] = None,
    ):
        """
        Inicializa a automação Selenium.
        
        Args:
            url: URL do site a automatizar
            headless: Se True, executa o navegador em modo headless
            progress_callback: Função para reportar progresso
        """
        self.url = url or config.SITE_URL
        self.progress_callback = progress_callback
        self.driver = None
        self.wait = None
        self._init_driver(headless)
    
    def _init_driver(self, headless: bool) -> None:
        """Inicializa o driver do Selenium."""
        options = webdriver.ChromeOptions()
        if headless:
            options.add_argument("--headless")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, config.ELEMENT_TIMEOUT)
    
    def _report_progress(self, message: str, current: int = 0, total: int = 0) -> None:
        """Reporta progresso via callback."""
        if self.progress_callback:
            self.progress_callback(message, current, total)
        logger.info(f"{message} ({current}/{total})" if total > 0 else message)
    
    def navigate(self) -> None:
        """Navega para o URL especificado."""
        self._report_progress(f"Acessando {self.url}")
        self.driver.get(self.url)
        time.sleep(2)
    
    def _fill_campo_texto(self, field_id: str, value: str) -> None:
        """Preenche um campo de texto."""
        try:
            element = self.wait.until(
                EC.presence_of_element_located((By.ID, field_id))
            )
            element.clear()
            element.send_keys(str(value))
            time.sleep(config.FILL_DELAY)
        except Exception as e:
            logger.error(f"Erro ao preencher campo {field_id}: {e}")
            raise
    
    def _select_categoria(self, categoria: str) -> None:
        """Seleciona uma categoria no dropdown."""
        try:
            select_element = self.wait.until(
                EC.presence_of_element_located((By.ID, "categoria"))
            )
            select = Select(select_element)
            select.select_by_value(categoria)
            time.sleep(config.FILL_DELAY)
        except Exception as e:
            logger.error(f"Erro ao selecionar categoria {categoria}: {e}")
            raise
    
    def _submit_form(self) -> None:
        """Submete o formulário."""
        try:
            btn_save = self.wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, config.BUTTONS["save"])
                )
            )
            btn_save.click()
            time.sleep(config.SUBMIT_DELAY)
        except Exception as e:
            logger.error(f"Erro ao submeter formulário: {e}")
            raise
    
    def _limpar_form(self) -> None:
        """Limpa o formulário."""
        try:
            btn_clear = self.driver.find_element(By.CSS_SELECTOR, config.BUTTONS["clear"])
            btn_clear.click()
            time.sleep(config.ACTION_DELAY)
        except Exception as e:
            logger.warning(f"Erro ao limpar formulário: {e}")
    
    def preencher_produto(self, product: Product) -> bool:
        """
        Preenche e submete um produto.
        
        Args:
            product: Objeto Product com dados a preencher
            
        Returns:
            True se sucesso, False caso contrário
        """
        try:
            self._limpar_form()
            self._fill_campo_texto("cliente", product.cliente)
            self._fill_campo_texto("produto", product.produto)
            self._fill_campo_texto("quantidade", product.quantidade)
            self._select_categoria(product.categoria)
            self._submit_form()
            return True
        except Exception as e:
            logger.error(f"Erro ao preencher produto {product.produto}: {e}")
            return False
    
    def processar_planilha(self, filepath: str) -> Dict[str, Any]:
        """
        Processa uma planilha Excel e preenche todos os produtos.
        
        Args:
            filepath: Caminho para o arquivo Excel
            
        Returns:
            Dicionário com estatísticas do processamento
        """
        stats = {
            "total": 0,
            "sucesso": 0,
            "erro": 0,
            "erros_detalhados": []
        }
        
        try:
            self.navigate()
            
            workbook = openpyxl.load_workbook(filepath)
            ws = workbook.active
            
            # Pula o cabeçalho
            rows = list(ws.iter_rows(values_only=True))[1:]
            stats["total"] = len(rows)
            
            self._report_progress(f"Iniciando processamento de {stats['total']} produtos")
            
            for index, row in enumerate(rows, 1):
                if not row or not row[0]:  # Pula linhas vazias
                    continue
                
                try:
                    cliente, produto, quantidade, categoria = row[:4]
                    
                    # Validação básica
                    if not all([cliente, produto, quantidade, categoria]):
                        raise ValueError("Campos obrigatórios vazios")
                    
                    product = Product(
                        cliente=str(cliente).strip(),
                        produto=str(produto).strip(),
                        quantidade=int(quantidade),
                        categoria=str(categoria).strip()
                    )
                    
                    self._report_progress(
                        f"Processando: {product.cliente} - {product.produto}",
                        index,
                        stats["total"]
                    )
                    
                    if self.preencher_produto(product):
                        stats["sucesso"] += 1
                    else:
                        stats["erro"] += 1
                        stats["erros_detalhados"].append({
                            "linha": index + 1,
                            "produto": produto,
                            "erro": "Falha ao submeter"
                        })
                
                except Exception as e:
                    stats["erro"] += 1
                    stats["erros_detalhados"].append({
                        "linha": index + 1,
                        "produto": row[1] if len(row) > 1 else "Desconhecido",
                        "erro": str(e)
                    })
                    logger.error(f"Erro na linha {index + 1}: {e}")
            
            self._report_progress(
                f"Processamento concluído: {stats['sucesso']} sucesso, {stats['erro']} erros"
            )
            
        except Exception as e:
            logger.error(f"Erro crítico ao processar planilha: {e}")
            stats["erros_detalhados"].append({
                "linha": 0,
                "produto": "ERRO CRÍTICO",
                "erro": str(e)
            })
        
        return stats
    
    def fechar(self) -> None:
        """Fecha o driver do Selenium."""
        if self.driver:
            self.driver.quit()
    
    def __enter__(self):
        """Context manager entry."""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.fechar()


def setup_logging(log_file: Optional[str] = None) -> None:
    """Configura o logging da aplicação."""
    log_format = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    
    logging.basicConfig(
        level=logging.INFO,
        format=log_format
    )
    
    if log_file:
        file_handler = logging.FileHandler(log_file)
        file_handler.setFormatter(logging.Formatter(log_format))
        logger.addHandler(file_handler)
