"""
Módulo para fornecer os blocos de CSS e HTML da aplicação.
"""

def get_css_block():
    """Retorna o bloco de CSS para estilizar a página."""
    return """
    <style>
        body {
            background-color: #FFFFFF;
        }
        .header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 10px;
            border-bottom: 2px solid #01a9e0;
        }
        .header h1 {
            color: #01a9e0;
            font-family: Arial, sans-serif;
            margin: 0;
        }
        .logo-container {
            display: flex;
            align-items: center;
            gap: 20px;
        }
        .logo {
            height: 60px;
        }
        .kpi {
            text-align: center;
            padding: 20px;
            margin: 10px;
            border-radius: 10px;
            background-color: #f9f9f9;
            border: 1px solid #ddd;
        }
        .kpi h2 {
            color: #01a9e0;
            font-size: 36px;
            margin-bottom: 10px;
        }
        .kpi p {
            font-size: 30px;
            color: #626366;
            font-weight: bold;
        }
    </style>
    """

def get_header_html():
    """Retorna o bloco de HTML para o cabeçalho."""
    return """
    <div class="header">
        <h1>Processador de Relatórios</h1>
        <div class="logo-container">
            <img src="https://i.postimg.cc/jjqqgPv6/logo-neo.png" alt="Logo da NEO" class="logo">
            <img src="https://docol65anos.com.br/temp/logo-docol.png" alt="Logo da Docol" class="logo">
        </div>
    </div>
    """