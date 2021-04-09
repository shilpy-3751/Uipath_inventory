Set IE = CreateObject("InternetExplorer.Application")
IE.Visible = True
IE.Navigate "https://opensource-demo.orangehrmlive.com/"

WScript.Sleep 10000

Set oUserNameCol = IE.Document.GetElementById("txtUsername")
oUserNameCol.Value = "Admin"

Set oUserPwdCol = IE.Document.GetElementById("txtPassword")
oUserPwdCol.Value = "admin123"

IE.Document.GetElementById("btnLogin").Click