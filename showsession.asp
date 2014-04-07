<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
  <%@ Language=VBScript %>
    <%
    on error resume next
    for x = 0 to session.Contents.Count
    Response.Write "<nobr><pre><b>"
    Response.write session.Contents.key(x) & " - </b>"
    Response.write session.Contents.item(x) & "</nobr><p><br>"
    next
    %>
</body>
</html>
