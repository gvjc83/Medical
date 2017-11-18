

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexion.asp"-->
<%

function LimpiarSQL(texto)
	dim tt
	tt=texto
	tt=replace(tt,"""","")
	tt=replace(tt,"'","")
	tt=replace(tt,"--","")
	tt=replace(tt,"insert","")
	tt=replace(tt,"delete","")
	tt=replace(tt,"drop","")
	tt=replace(tt,"truncate","")
	tt=replace(tt,"@@","")
	LimpiarSQL=tt
end function


op = LimpiarSQL(request("op"))

if op = 1 then

	rut = LimpiarSQL(request("rut"))
	pass = LimpiarSQL(request("pass"))

	Set rs = Server.CreateObject("ADODB.RecordSet")
	sql = "select id_usuario,u_usuario,id_cliente, nombres, a_paterno, email from usuario where estado = 1 and rut = '"&rut&"' and password = '"&pass&"'"
	rs.Open sql,conn

	IF NOT rs.EOF THEN
		session("id_usuario") = rs("id_usuario")
		session("u_usuario")  = rs("u_usuario")
		session("id_cliente") = rs("id_cliente")
		session("nombres")    = rs("nombres")
		session("a_paterno")  = rs("a_paterno")
		session("email")      = rs("email")
		
		Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
		Response.Write("<login>"&chr(13)) 
		Response.Write("<row>"&chr(13))
		
		Response.Write("<URL>1</URL>"&chr(13))
		Response.Write("<HTTP>inicio.asp</HTTP>"&chr(13))
		
		Response.Write("</row>"&chr(13))
		Response.Write("</login>") 		
		
	ELSE
		session("id_usuario") = ""
		session("u_usuario")  = ""
		session("id_cliente") = ""
		session("nombres")    = ""
		session("a_paterno")  = ""
		session("email")      = ""
		
		Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
		Response.Write("<login>"&chr(13)) 
		Response.Write("<row>"&chr(13))
		
		Response.Write("<URL>0</URL>"&chr(13))
		Response.Write("<HTTP></HTTP>"&chr(13))
		
		Response.Write("</row>"&chr(13))
		Response.Write("</login>") 				
	END IF
			
	Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
	Response.Write("<login>"&chr(13)) 
	
	WHILE NOT rs.EOF
			Response.Write("<row>"&chr(13))
			Response.Write("<id_usuario>"&rs("id_usuario")&"</ID>"&chr(13))
			Response.Write("<u_usuario>"&rs("u_usuario")&"</DESC>"&chr(13))
			Response.Write("<id_cliente>"&rs("id_cliente")&"</ID>"&chr(13))
			Response.Write("<nombres>"&rs("nombres")&"</DESC>"&chr(13))
			Response.Write("<a_paterno>"&rs("a_paterno")&"</ID>"&chr(13))
			Response.Write("<email>"&rs("email")&"</DESC>"&chr(13))
			Response.Write("</row>"&chr(13))
		rs.MoveNext
	WEND
	Response.Write("</login>") 
end if'op=1
%>