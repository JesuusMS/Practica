<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>


<%
Dim MM_servidor_STRING
MM_servidor_STRING = "PROVIDER=SQLOLEDB;DATA SOURCE=LAPTOP-297K79P6\TLX_JAAG;UID=sa;PWD=12345;DATABASE=CMJA"
%>

<%
nombre=request.form("nombre")
tel=request.form("tel")
%>


<%
set Inserta_tel = Server.CreateObject("ADODB.Command")
Inserta_tel.ActiveConnection = MM_servidor_STRING
Inserta_tel.CommandText = "INSERT INTO   agenda_jm (nombre,tel) VALUES('" + Replace(nombre, "'", "''") + "','" + Replace(tel, "'", "''") + "')"
'response.write Inserta_tel.CommandText
Inserta_tel.CommandType = 1
Inserta_tel.CommandTimeout = 0
Inserta_tel.Prepared = true
Inserta_tel.Execute()
%>


<head>
<title>Agenda</title>
</head>

<body>
Se guardo la informaci&oacute;n correctamente
</body>