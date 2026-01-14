import smtplib
import pandas as pd
import time
import random
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from google.colab import drive

# 1. CONECTAR AO DRIVE
drive.mount('/content/drive', force_remount=True)

def enviar_ofertas_colab():
    # --- CONFIGURAÇÕES ---
    meu_email = "XXXXXXXXXXX@gmail.com"
    minha_senha = "XXXXXXXXXX"
    meu_whatsapp = "https://wa.me/5519XXXXXX"
    link_site = "https://isa-duque.github.io/portal-cnpj-mei/"
    
    caminho_arquivo = "/content/drive/MyDrive/LEADS_EXTRACAO_20260101.xlsx"
    
    # --- PONTO DE RETOMADA ---
    ultimo_email_enviado = "marmitariadodon@gmail.com"
    encontrou_ponto_parada = False 

    try:
        dados = pd.read_excel(caminho_arquivo)
        print(f"Planilha carregada. Total de linhas: {len(dados)}")
        
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(meu_email, minha_senha)

        contador_envios = 0

        for i, linha in dados.iterrows():
            email_cliente = str(linha['email']).strip().lower()
            cnpj_cliente = str(linha['cnpj_completo']).strip()

            # LÓGICA DE RETOMADA
            if not encontrou_ponto_parada:
                if email_cliente == ultimo_email_enviado.lower():
                    encontrou_ponto_parada = True
                    print(f"Ponto de retomada encontrado: {email_cliente}. Iniciando envios...")
                continue 

            # MONTAGEM DO CORPO DO EMAIL
            msg = MIMEMultipart()
            msg['From'] = meu_email
            msg['To'] = email_cliente
            msg['Subject'] = f"Maquininha e Crédito para seu novo CNPJ ! ({cnpj_cliente})"

            corpo_html = f"""
            <html>
                <body>
                    <p>Olá!</p>
                    <p>Vi que vocês iniciaram as atividades recentemente (CNPJ: {cnpj_cliente}) e passei para desejar <strong>Muito SUCESSO!</strong></p>
                    <p>Sou especialista em serviços financeiros e ajudo novas empresas a garantirem as <strong>melhores taxas de maquininha</strong> do mercado e acesso a <strong>crédito facilitado</strong> para capital de giro.</p>
                    <p>Maquininha <strong>sem ALUGUEL</strong>, Tenho cupom de <strong>82% DE DESCONTO</strong> na aquisição da maquininha e taxa Promocional de 0,53%*</p>
                    <p><a href="{meu_whatsapp}" style="background-color: #25D366; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-weight: bold;">Falar com Consultor via WhatsApp</a></p>
                    <br>
                    <p>Grande abraço,<br><strong>Isabela Ferreira</strong><br>
                    SITE: <a href="{link_site}">{link_site}</a></p>
                </body>
            </html>
            """
            
            msg.attach(MIMEText(corpo_html, 'html'))
            server.send_message(msg)
            
            contador_envios += 1
            print(f"[{contador_envios}/400] Enviado para: {email_cliente}")

            if contador_envios >= 400:
                print("Limite de 400 envios atingido.")
                break

            # Pausa humana (15 a 25 segundos)
            time.sleep(random.randint(15, 25))

        server.quit()
        print("\n--- Finalizado com sucesso! ---")

    except Exception as e:
        print(f"\nERRO: {e}")

if __name__ == "__main__":
    enviar_ofertas_colab()
