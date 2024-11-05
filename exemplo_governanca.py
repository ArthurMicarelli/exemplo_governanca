import logging
import win32com.client as win32

# Configurando o logger para o Jupyter Notebook
logger = logging.getLogger()
logger.setLevel(logging.INFO)

# Configurando o logger para exibir no output do notebook
stream_handler = logging.StreamHandler()
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
stream_handler.setFormatter(formatter)
if not logger.hasHandlers():
    logger.addHandler(stream_handler)

# Função para enviar e-mail em caso de erro usando Outlook
def send_error_email(error_message):
    try:
        # Configuração do e-mail
        receiver_email = "arthur.domingos@germinare.org.br"  # E-mail do destinatário
        subject = "Erro no Script Python"
        body = f"Ocorreu o seguinte erro no seu script:\n\n{error_message}"

        # Enviando o e-mail com Outlook via win32com
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = receiver_email
        mail.Subject = subject
        mail.Body = body
        mail.Send()
        
        logger.info("E-mail de erro enviado com sucesso via Outlook.")
    except Exception as e:
        logger.error(f"Falha ao enviar o e-mail: {e}")

# Função principal para simular um erro
def main():
    try:
        logger.info("Iniciando o script...")
        
        # Simulação de erro: divisão por zero
        result = 10 / 0
        
        logger.info(f"Resultado da operação: {result}")
    except Exception as e:
        error_message = str(e)
        logger.error(f"Ocorreu um erro: {error_message}")
        send_error_email(error_message)

# Executando o script
main()
