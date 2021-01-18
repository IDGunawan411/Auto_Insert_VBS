Set oConDW = CreateObject("ADODB.Connection")
oConDW.ConnectionString = "Password=password;User ID=reportintra;Data Source=reportintra"
oConDW.Open()
wscript.echo "connection status :" & oConDW