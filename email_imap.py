import imaplib
import email
import os
import json
from datetime import datetime, timedelta
import time
import subprocess
import platform
from pathlib import Path
import threading
import keyboard
from dateutil import tz

# Esse programa baixa anexos do Gmail através de uma conexão IMAP
# usando login com email e senha de chave de app gerada nas configuraçõesdo Google.
# Foi necessario gerar a chave de app para cada conta. Para gerar a chave de app é necessaário
# que a conta tenha verificação em 2 etapas ativada.
# O programa baixa anexos de qualquer serviço de email, desde que especificado o servidor imap

user = Path.home()
CONFIG_FILE = "config.json"
DEFAULT_SAVE_DIR = str(user) + "\Anexos"
desired_date = None  # Variável global para armazenar o diretório selecionado anteriormente

def convert_to_brasilia_timezone(original_datetime):
    original_tz = tz.tzoffset(None, -8 * 60 * 60)  # Fuso horário original do e-mail (-08:00)
    brasilia_tz = tz.gettz("America/Sao_Paulo")
    original_datetime = original_datetime.replace(tzinfo=original_tz)
    brasilia_datetime = original_datetime.astimezone(brasilia_tz)
    return brasilia_datetime

def download_attachments(mail, msg, save_attachments_dir):
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        filename = part.get_filename()
        if not filename:
            continue

        original_filename = filename

        email_datetime = get_email_datetime(msg)
        if email_datetime:
            timestamp = convert_to_brasilia_timezone(email_datetime)  # Use o objeto datetime diretamente
        else:
            timestamp = datetime.now()

        timestamp_str = timestamp.strftime("%Y-%m-%d_%H-%M-%S")
        filename_with_timestamp = f"{timestamp_str}_{filename}"
        filepath = os.path.join(save_attachments_dir, filename_with_timestamp)

        with open(filepath, 'wb') as f:
            f.write(part.get_payload(decode=True))


def open_folder_in_explorer(PREVIOUS_SAVE_DIR):
    system = platform.system()
    if system == "Windows":
        os.startfile(PREVIOUS_SAVE_DIR)
    elif system == "Darwin":
        subprocess.run(["open", PREVIOUS_SAVE_DIR])
    elif system == "Linux":
        subprocess.run(["xdg-open", PREVIOUS_SAVE_DIR])

def ensure_directory_exists(directory_path):
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)

def select_save_dir():
    global PREVIOUS_SAVE_DIR
    print(f"Diretório padrão pré-selecionado: {DEFAULT_SAVE_DIR}")
    choice = input("Deseja usar o diretório padrão? (s/n): ")
    if choice.lower() == "s":
        PREVIOUS_SAVE_DIR = DEFAULT_SAVE_DIR
        return DEFAULT_SAVE_DIR
    else:
        while True:
            save_attachments_dir = input("Digite o diretório onde deseja salvar os anexos: ")
            if not save_attachments_dir:
                print("Diretório inválido. Tente novamente.")
            else:
                if not os.path.exists(save_attachments_dir):
                    os.makedirs(save_attachments_dir)
                PREVIOUS_SAVE_DIR = save_attachments_dir
                return save_attachments_dir

def read_config():
    with open(CONFIG_FILE, "r") as f:
        return json.load(f)

def monitor_keyboard_input():

    # Monitora a entrada de teclado usando a biblioteca 'keyboard'
    while True:
        try:
            if keyboard.is_pressed("ctrl+o"):
                print(f"Abrindo o diretório: {PREVIOUS_SAVE_DIR}")
                open_folder_in_explorer(PREVIOUS_SAVE_DIR)
                while keyboard.is_pressed("ctrl+o"):
                    time.sleep(0.1)
        except:
            # Caso ocorra algum erro (por exemplo, o terminal não suporta essa funcionalidade), continua o loop
            print("Erro ao abrir o caminho da pasta")
            pass


def get_email_datetime(msg):
    date_str = msg["Date"]
    try:
        date_obj = email.utils.parsedate_to_datetime(date_str)
        print(date_obj)
        return date_obj
    except:
        return None

def select_account(accounts):
    print("Contas disponíveis:")
    for i, account in enumerate(accounts, 1):
        print(f"{i}. {account['email']}")

    while True:
        try:
            choice = int(input("Escolha uma conta (digite o número correspondente): "))
            if 1 <= choice <= len(accounts):
                return accounts[choice - 1]
            else:
                print("Opção inválida. Tente novamente.")
        except ValueError:
            print("Entrada inválida. Digite um número.")
            
def auto_mode_select():
    return False if input("Deseja escolher manualmente a data em que os e-mails serão baixados? (Digite 's' para sim "
                          "ou deixe em branco para usar a data de hoje: ").lower() == "s" else True

def get_desired_date(auto_mode=True):
    if not auto_mode:
        while True:
            try:
                current_date_str = datetime.now().strftime("%d/%m/%Y")
                date_str = input(f"Digite a data desejada no formato DD/MM/AAAA (ou pressione Enter para usar a data "
                                 f"atual {current_date_str}): ")
                if not date_str:
                    return datetime.now().date()

                desired_date = datetime.strptime(date_str, "%d/%m/%Y").date()
                return desired_date
            except ValueError:
                print("Data inválida. Tente novamente.")
    else:
        return datetime.now().date()

def connect_to_gmail(email, password):
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(email, password)
    mail.select("inbox")
    return mail

def mark_as_read(mail, msg_id):
    mail.store(msg_id, '+FLAGS', '\Seen')

def mark_as_unread(mail, msg_id):
    mail.store(msg_id, '-FLAGS', '\Seen')
    unread = mail.store(msg_id, '-FLAGS', '\Seen')
    print("Retorno da função mark_as_unread", unread)

def main():
    # Inicia a thread para monitorar a entrada de teclado em segundo plano
    keyboard_thread = threading.Thread(target=monitor_keyboard_input)
    keyboard_thread.daemon = True
    keyboard_thread.start()

    auto_mode = auto_mode_select()

    while True:
        try:
            config = read_config()
            account = select_account(config["accounts"])

            email_address = account["email"]
            email_password = account["password"]

            mail = connect_to_gmail(email_address, email_password)

            while True:
                order = input(
                    "Digite 'novo' para baixar do mais novo para o mais antigo, ou 'antigo' para baixar do mais antigo "
                    "para o mais novo: "
                )
                if order.lower() in ["novo", "antigo"]:
                    break
                else:
                    print("Opção inválida. Tente novamente.")

            save_attachments_dir = select_save_dir()

            while True:
                desired_date = get_desired_date(auto_mode=auto_mode)

                ensure_directory_exists(DEFAULT_SAVE_DIR)

                end_date = desired_date + timedelta(days=1)
                date_format = "%d-%b-%Y"
                since_date_str = desired_date.strftime(date_format)
                end_date_str = end_date.strftime(date_format)

                search_criteria = f'(UNSEEN SINCE "{since_date_str}" BEFORE "{end_date_str}")'

                result, data = mail.search(None, search_criteria)
                ids = data[0].split()

                if order.lower() == "novo":
                    ids = ids[::-1]

                for i in ids:
                    print("ids: ", ids)
                    try:
                        result, data = mail.fetch(i, "(RFC822)")
                        if data[0] is not None and isinstance(data[0][1],
                                                              bytes):  # Verifica se data[0][1] é um objeto de bytes
                            raw_email = data[0][1]
                            msg = email.message_from_bytes(raw_email)
                            ensure_directory_exists(save_attachments_dir)
                            try:
                               download_attachments(mail, msg, save_attachments_dir)
                               mark_as_read(mail, i)
                            except Exception as e:
                                mark_as_unread(mail, i)
                                print("Erro no download, tentando marcar como não lido. Err: ", e
                                      , " . Anexo: ", i)
                            print(f"Anexo {i} baixado com sucesso")

                    except Exception as e:
                        t = connect_to_gmail(email_address, email_password)
                        print("Tentando reconexão com gmail: ", t)
                        mark_as_unread(mail, i)
                        print(f"Erro ao baixar anexo {i}")
                        print(f"Erro durante o download do anexo: {e}")

                    # Adiciona uma pausa entre o download de anexos para evitar problemas de conexão
                    time.sleep(2)

                mail.expunge()


                interval = 60
                print(f"Aguardando {interval} segundos...")
                time.sleep(interval)


        except Exception as erro:
            mark_as_unread(mail, i)
            print(f"Email marcado como 'não lido': mail: {mail}, i: {i}")
            print(f"Houve um erro: {erro}")
            input("Pressione Enter")

        except KeyboardInterrupt:
            print("Pressione Ctrl+R para reiniciar o programa.")
            while True:
                if input() == "\x12":
                    break

if __name__ == "__main__":
    main()
    # Implementar uma função full auto que vai alternando entre contas, sem eu precisar ficar vindo aqui trocar
    # a conta manualmente.
    # Entender melhor esse código, corrigir, otimizar, excluir o que não for necessário, e melhorar a estrutrura
    # dele.
    # Fazer ele ser compativel com outros tipos de email, já que uso o protocolo SMTP com ele ao invés de uma API
    # de algum serviço de email especifico. Eu posso perguntar para o usuário, ou então identificar o tipo de
    # email e automaticamente colocar a porta smtp dele. A do google é: mail = imaplib.IMAP4_SSL("imap.gmail.com")
