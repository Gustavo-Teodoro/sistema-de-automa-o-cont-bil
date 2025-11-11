"""
Script para testar a automação sem a GUI (modo console).
Útil para debug e testes.
"""

import sys
import logging
from pathlib import Path

from selenium_automation import SeleniumAutomation, setup_logging


def progress_callback(message: str, current: int = 0, total: int = 0) -> None:
    """Callback para reportar progresso."""
    if total > 0:
        percentage = (current / total) * 100
        print(f"\r[{'='*40}] {percentage:.1f}% - {message}", end="", flush=True)
    else:
        print(f"✓ {message}")


def main():
    """Função principal para teste."""
    setup_logging("automacao.log")
    
    # Caminho da planilha
    excel_file = Path(__file__).parent / "dados.xlsx"
    
    if not excel_file.exists():
        print(f"❌ Arquivo não encontrado: {excel_file}")
        sys.exit(1)
    
    print("="*50)
    print("🚀 Teste de Automação Selenium")
    print("="*50)
    print(f"📁 Arquivo: {excel_file}")
    
    try:
        with SeleniumAutomation(
            headless=False,
            progress_callback=progress_callback
        ) as automation:
            stats = automation.processar_planilha(str(excel_file))
            
            print("\n" + "="*50)
            print("✓ Processamento Concluído")
            print("="*50)
            print(f"Total: {stats['total']}")
            print(f"✓ Sucessos: {stats['sucesso']}")
            print(f"✗ Erros: {stats['erro']}")
            
            if stats['erros_detalhados']:
                print("\n📋 Erros Detalhados:")
                for erro in stats['erros_detalhados']:
                    print(f"  Linha {erro['linha']}: {erro['erro']}")
    
    except Exception as e:
        print(f"\n❌ Erro: {e}")
        logging.exception(e)
        sys.exit(1)


if __name__ == "__main__":
    main()
