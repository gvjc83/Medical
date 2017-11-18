<%
public conn 
Set conn = Server.CreateObject("ADODB.Connection") 
conn.CommandTimeout = 0
conn.Open("Provider=SQLOLEDB; User ID=usrMedical;Password=Medical,.0;data Source=.\SQLEXPRESS;Initial Catalog=admin_medical") 
%>