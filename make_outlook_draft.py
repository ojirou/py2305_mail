import win32com.client
outlook = win32com.client.Dispatch('Outlook.Application')
mymail = outlook.CreateItem(0)
mymail.BodyFormat = 1
mymail.To = 'aaa@hoge.co.jp; bbb@hoge.co.jp'
mymail.cc = 'ccc@hoge.com'
mymail.Bcc = 'ddd@hoge.com'
mymail.Subject = 'たいとる'
mymail.Body = '''
お疲れ様です

以上、よろしくお願いします
mymail.Display(True)
#mymail.Send()
