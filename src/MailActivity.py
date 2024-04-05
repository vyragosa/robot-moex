import win32com.client as win32


def send_mail(to: str, subject: str, body: str, attachment: str) -> None:
    """
    Отправляет письмо через Outlook
    :param to: Кому отправить письмо
    :param subject: Тема письма
    :param body: Содержимое письма
    :param attachment: Вложения письма
    """
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(attachment)
    mail.Send()
