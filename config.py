"""
Configurações da aplicação.
"""

# URL do site a automatizar
SITE_URL = "https://devaprender-contabil.netlify.app/"

# Modo headless (True = sem interface do navegador)
HEADLESS_MODE = False

# Timeout para esperar elementos (segundos)
ELEMENT_TIMEOUT = 10

# Delay entre ações (segundos)
ACTION_DELAY = 0.5

# Delay entre preenchimentos (segundos)
FILL_DELAY = 0.3

# Delay após submit (segundos)
SUBMIT_DELAY = 1.0

# Arquivo de log
LOG_FILE = "automacao.log"

# Categorias disponíveis no site
VALID_CATEGORIES = {
    "Eletrônicos",
    "Móveis",
    "Roupas",
    "Brinquedos",
    "Comida",
    "Bebidas",
    "Cosméticos",
    "Livros",
    "Esportes",
    "Jardinagem"
}

# IDs dos campos HTML (baseado em home.html)
FORM_FIELDS = {
    "cliente": "cliente",
    "produto": "produto",
    "quantidade": "quantidade",
    "categoria": "categoria"
}

# Botões
BUTTONS = {
    "save": "button.btn-save",
    "clear": "button.btn-clear"
}
