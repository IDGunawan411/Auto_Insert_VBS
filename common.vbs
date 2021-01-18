
    ' Set MyEmail=CreateObject("CDO.Message")
    ' MyEmail.Subject="Subject"
    ' MyEmail.From="ssl://gunawanprasetyo313@gmail.com"
    ' MyEmail.To="ssl://reclosher@gmail.com"
    ' MyEmail.TextBody="Testing one two three."
    ' MyEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")= 2
    ' 'SMTP Server
    ' ' MyEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
    ' MyEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="smtp.gmail.com"
    ' 'SMTP Port
    ' MyEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")= 587
    ' MyEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "gunawanprasetyo313@gmail.com"
    ' MyEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "gunawan12345"
    ' MyEmail.Configuration.Fields.Update
    ' MyEmail.Send
    ' set MyEmail=nothing
    ' schema = "http://schemas.microsoft.com/cdo/configuration/"

    ' Set msg = CreateObject("CDO.Message")
    ' msg.Subject  = "Test"
    ' msg.From     = "gunawanprasetyo313@gmail.com"
    ' msg.To       = "reclosher@gmail.com"
    ' msg.TextBody = "This is some sample message text."

    ' With msg.Configuration.Fields
    '     .Item(schema & "sendusing")      = 2
    '     .Item(schema & "smtpserver")     = "smtp.gmail.com"
    '     .Item(schema & "smtpserverport") = 25
    '     .Update
    ' End With

    ' msg.Send
