REM  *****  BASIC  *****

Sub SendEmail
	Dim eMailAddress as String
	Dim eSubject as String
	Dim eMailer as Object
	Dim eMailClient as Object
	Dim eMessage as Object
	
	eMailAddress = "icigiligici69@gmail.com"
	eSubject = "Spreadsheet email"
	eMailer = createUnoService("com.sun.star.system.SimpleSystemMail")
	
	eMailClient = eMailer.querySimpleMailClient()
	
	eMessage = eMailClient.createSimpleMailMessage()
	eMessage.body = GetBody()
	eMessage.setRecipient(eMailAddress)
	eMessage.setSubject(eSubject)
	
	eMessage.setAttachement(Array(convertToUrl(GetCurrentFile())))
	
	eMailClient.sendSimpleMailMessage(eMessage, com.sun.star.system.SimpleMailClientFlags.NO_USER_INTERFACE)
End Sub

Function GetCurrentFile
	Dim path as String
	path = ThisComponent.getURL()
	GetCurrentFile = path
End Function


Function GetBody()
	GetBody = InputBox("Enter the body of the email", "Input Box")
End Function