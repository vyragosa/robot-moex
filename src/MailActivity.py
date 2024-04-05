import win32com.client as win32


def send_mail(to: str, subject: str, body: str, attachment: str) -> None:
    """
    Отправляет письмо через Outlook
    :param to: Кому отправить письмо
    :param subject: Тема письма
    :param body: Содержимое письма
    :param attachment: Вложения письма
    """
    print("Opening outlook...")
    outlook = win32.Dispatch('outlook.application') # могут быть ошибки с запуском. Зависит от версии outlook.
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(attachment)
    mail.Send()
