import threading
import logging
from pathlib import Path
from typing import Optional

import customtkinter as ctk
from tkinter import filedialog, messagebox, scrolledtext

from selenium_automation import SeleniumAutomation, setup_logging


logger = logging.getLogger(__name__)


class AutomationGUI(ctk.CTk):
    """GUI principal para a automação."""
    
    def __init__(self):
        """Inicializa a janela principal."""
        super().__init__()
        
        # Configurações da janela
        self.title("Automação - Sistema de Contabilidade")
        self.geometry("900x700")
        self.minsize(800, 600)
        
        # Tema
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        # Variáveis
        self.file_path: Optional[str] = None
        self.automation: Optional[SeleniumAutomation] = None
        self.is_running = False
        
        # Setup logging
        setup_logging()
        
        # Criar interface
        self._create_widgets()
    
    def _create_widgets(self) -> None:
        """Cria todos os widgets da interface."""
        # Frame principal
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Header
        self._create_header(main_frame)
        
        # Frame de seleção de arquivo
        self._create_file_section(main_frame)
        
        # Frame de progresso
        self._create_progress_section(main_frame)
        
        # Frame de log
        self._create_log_section(main_frame)
        
        # Frame de botões
        self._create_button_section(main_frame)
    
    def _create_header(self, parent) -> None:
        """Cria o header com título."""
        header_frame = ctk.CTkFrame(parent, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 20))
        
        title = ctk.CTkLabel(
            header_frame,
            text="📊 Automação de Cadastro de Produtos",
            font=("Segoe UI", 24, "bold")
        )
        title.pack(anchor="w")
        
        subtitle = ctk.CTkLabel(
            header_frame,
            text="Envie sua planilha e automatize o cadastro de produtos",
            font=("Segoe UI", 12),
            text_color="gray"
        )
        subtitle.pack(anchor="w")
    
    def _create_file_section(self, parent) -> None:
        """Cria a seção de seleção de arquivo."""
        file_frame = ctk.CTkFrame(parent)
        file_frame.pack(fill="x", pady=(0, 15))
        
        # Label
        label = ctk.CTkLabel(
            file_frame,
            text="📁 Selecione a Planilha Excel",
            font=("Segoe UI", 12, "bold")
        )
        label.pack(anchor="w", pady=(0, 10))
        
        # Frame para botão e caminho
        path_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        path_frame.pack(fill="x")
        
        self.file_path_label = ctk.CTkLabel(
            path_frame,
            text="Nenhum arquivo selecionado",
            font=("Segoe UI", 10),
            text_color="gray"
        )
        self.file_path_label.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        btn_select = ctk.CTkButton(
            path_frame,
            text="Selecionar Arquivo",
            command=self._select_file,
            width=150,
            fg_color="#667eea",
            hover_color="#764ba2"
        )
        btn_select.pack(side="right")
    
    def _create_progress_section(self, parent) -> None:
        """Cria a seção de progresso."""
        progress_frame = ctk.CTkFrame(parent)
        progress_frame.pack(fill="x", pady=(0, 15))
        
        # Label
        label = ctk.CTkLabel(
            progress_frame,
            text="📈 Progresso",
            font=("Segoe UI", 12, "bold")
        )
        label.pack(anchor="w", pady=(0, 10))
        
        # Status
        self.status_label = ctk.CTkLabel(
            progress_frame,
            text="Aguardando início...",
            font=("Segoe UI", 10),
            text_color="gray"
        )
        self.status_label.pack(anchor="w", pady=(0, 8))
        
        # Progresso
        self.progress_bar = ctk.CTkProgressBar(
            progress_frame,
            fg_color="gray30",
            progress_color="#667eea"
        )
        self.progress_bar.pack(fill="x", pady=(0, 8))
        self.progress_bar.set(0)
        
        # Informações
        info_frame = ctk.CTkFrame(progress_frame, fg_color="transparent")
        info_frame.pack(fill="x")
        
        self.progress_info = ctk.CTkLabel(
            info_frame,
            text="0 / 0 processados",
            font=("Segoe UI", 9),
            text_color="gray"
        )
        self.progress_info.pack(side="left")
        
        self.stats_label = ctk.CTkLabel(
            info_frame,
            text="",
            font=("Segoe UI", 9),
            text_color="gray"
        )
        self.stats_label.pack(side="right")
    
    def _create_log_section(self, parent) -> None:
        """Cria a seção de log."""
        log_frame = ctk.CTkFrame(parent)
        log_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        # Label
        label = ctk.CTkLabel(
            log_frame,
            text="📋 Log de Atividades",
            font=("Segoe UI", 12, "bold")
        )
        label.pack(anchor="w", pady=(0, 10))
        
        # Text widget
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=12,
            bg="#2b2b2b",
            fg="#ffffff",
            insertbackground="white",
            wrap="word",
            font=("Consolas", 9)
        )
        self.log_text.pack(fill="both", expand=True)
        
        # Configurar tags de cores
        self.log_text.tag_config("info", foreground="#87CEEB")
        self.log_text.tag_config("success", foreground="#90EE90")
        self.log_text.tag_config("error", foreground="#FF6B6B")
        self.log_text.tag_config("warning", foreground="#FFD700")
    
    def _create_button_section(self, parent) -> None:
        """Cria a seção de botões."""
        button_frame = ctk.CTkFrame(parent, fg_color="transparent")
        button_frame.pack(fill="x", pady=(0, 0))
        
        self.btn_start = ctk.CTkButton(
            button_frame,
            text="▶️ Iniciar Automação",
            command=self._start_automation,
            fg_color="#667eea",
            hover_color="#764ba2",
            height=40,
            font=("Segoe UI", 12, "bold")
        )
        self.btn_start.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        self.btn_stop = ctk.CTkButton(
            button_frame,
            text="⏹️ Parar",
            command=self._stop_automation,
            fg_color="#ff6b6b",
            hover_color="#ff5252",
            height=40,
            font=("Segoe UI", 12, "bold"),
            state="disabled"
        )
        self.btn_stop.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        btn_clear_log = ctk.CTkButton(
            button_frame,
            text="🗑️ Limpar Log",
            command=self._clear_log,
            fg_color="gray40",
            hover_color="gray50",
            height=40,
            font=("Segoe UI", 12, "bold")
        )
        btn_clear_log.pack(side="left", fill="x", expand=True)
    
    def _select_file(self) -> None:
        """Abre diálogo para selecionar arquivo."""
        file_path = filedialog.askopenfilename(
            title="Selecione a planilha Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            initialdir=str(Path.home())
        )
        
        if file_path:
            self.file_path = file_path
            file_name = Path(file_path).name
            self.file_path_label.configure(text=file_name, text_color="white")
            self._log("✓ Arquivo selecionado: " + file_name, "success")
    
    def _start_automation(self) -> None:
        """Inicia a automação em thread separada."""
        if not self.file_path:
            messagebox.showwarning(
                "Arquivo não selecionado",
                "Por favor, selecione uma planilha Excel primeiro."
            )
            return
        
        if not Path(self.file_path).exists():
            messagebox.showerror(
                "Arquivo não encontrado",
                f"O arquivo {self.file_path} não existe."
            )
            return
        
        self.is_running = True
        self._update_ui_running(True)
        
        # Iniciar automação em thread separada
        thread = threading.Thread(target=self._run_automation, daemon=True)
        thread.start()
    
    def _run_automation(self) -> None:
        """Executa a automação."""
        try:
            self._log("🚀 Iniciando automação...", "info")
            
            automation = SeleniumAutomation(
                headless=False,
                progress_callback=self._update_progress
            )
            self.automation = automation
            
            try:
                stats = automation.processar_planilha(self.file_path)
                
                # Log dos resultados
                self._log("\n" + "="*50, "info")
                self._log("✓ Automação Concluída!", "success")
                self._log("="*50, "info")
                self._log(f"Total processado: {stats['total']}", "info")
                self._log(f"✓ Sucessos: {stats['sucesso']}", "success")
                self._log(f"✗ Erros: {stats['erro']}", "error" if stats['erro'] > 0 else "success")
                
                if stats['erros_detalhados']:
                    self._log("\n📋 Detalhes dos erros:", "warning")
                    for erro in stats['erros_detalhados']:
                        self._log(
                            f"  Linha {erro['linha']}: {erro['produto']} - {erro['erro']}",
                            "error"
                        )
                
                messagebox.showinfo(
                    "Sucesso",
                    f"Automação concluída!\n\n"
                    f"Processados: {stats['total']}\n"
                    f"Sucessos: {stats['sucesso']}\n"
                    f"Erros: {stats['erro']}"
                )
            
            finally:
                automation.fechar()
                self.automation = None
        
        except Exception as e:
            error_msg = f"Erro durante automação: {str(e)}"
            self._log(error_msg, "error")
            messagebox.showerror("Erro", error_msg)
            logger.exception(e)
        
        finally:
            self.is_running = False
            self._update_ui_running(False)
    
    def _stop_automation(self) -> None:
        """Para a automação."""
        if self.automation:
            self._log("⏹️ Parando automação...", "warning")
            self.automation.fechar()
            self.automation = None
            self.is_running = False
            self._update_ui_running(False)
    
    def _update_progress(
        self,
        message: str,
        current: int = 0,
        total: int = 0
    ) -> None:
        """Atualiza a barra de progresso."""
        self._log(message, "info")
        
        if total > 0:
            progress = current / total
            self.progress_bar.set(progress)
            self.progress_info.configure(text=f"{current} / {total} processados")
            self.status_label.configure(text=message)
            
            if current > 0:
                self.stats_label.configure(text=f"{progress*100:.1f}%")
    
    def _log(self, message: str, tag: str = "info") -> None:
        """Adiciona mensagem ao log."""
        self.log_text.insert("end", message + "\n", tag)
        self.log_text.see("end")
        self.update()
    
    def _clear_log(self) -> None:
        """Limpa o log."""
        self.log_text.delete("1.0", "end")
        self._log("Log limpo.", "info")
    
    def _update_ui_running(self, running: bool) -> None:
        """Atualiza estado dos botões."""
        if running:
            self.btn_start.configure(state="disabled")
            self.btn_stop.configure(state="normal")
            self.file_path_label.configure(state="disabled")
            self.progress_bar.set(0)
            self._log("\n🔄 Automação iniciada...", "info")
        else:
            self.btn_start.configure(state="normal")
            self.btn_stop.configure(state="disabled")
            self.file_path_label.configure(state="normal")
            self._log("✓ Automação finalizada.", "success")


def main():
    """Função principal."""
    app = AutomationGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
