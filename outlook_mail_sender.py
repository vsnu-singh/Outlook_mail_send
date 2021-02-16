import win32com.client as win32

class mail:
    def __init__(self):
        pass
    def send(mail_to,message="",subject="",attachment=""):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = mail_to
        mail.Subject = subject
        mail.Body = message
        mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

        # To attach a file to the email (optional):
        if(attachment!=""):
            try:
                mail.Attachments.Add(attachment)
            except:
                raise Exception("Invalid address of attachments")


        try:
            mail.Send()
        except:
            raise Exception("invalid email address")
    
