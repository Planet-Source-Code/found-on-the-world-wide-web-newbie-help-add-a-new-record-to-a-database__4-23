<div align="center">

## Newbie help: Add a new record to a database


</div>

### Description

Simple form to insert a value into a table using ADO's RecordSet.AddNew.

Name this form:addnew.ASP
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Beginner
**User Rating**    |3.0 (12 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-newbie-help-add-a-new-record-to-a-database__4-23/archive/master.zip)





### Source Code

```
<% Response.Expires = 0 %>
<HTML>
<BODY BGColor=Black Text=White>
<%
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application("guestDSN")
MySQL = "SELECT * FROM paulen"
Set RS = Server.CreateObject("ADODB.Recordset")
rs.Open MySQL, Conn, adOpenStatic, adLockOptimistic
rs.AddNew
rs(0) = Request.QueryString("fld1")
rs(1) = CInt(Request.QueryString("fld2"))
' rs(2) = Request.QueryString("fld3") ' field 3 is a text field, update not supported
rs.Update
%>
<B>New Record:</b><BR>
<%= rs(0) %>
<%= rs(1) %>
</BODY>
</HTML>
```

