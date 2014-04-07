<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Detect Smart Image Processor Components</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
function DetectImageComponent(DotNetResize)
  Dim objPictureProcessor, objASPjpeg, AspImage, AspSmart, objImgWriter, objAspThumb, ImageComponent, NetImageComponent
  ImageComponent = ""
  NetImageComponent = ""
  on error resume next
 'Check for our own Picture Processor
  err.clear
  Set objPictureProcessor = Server.CreateObject("COMobjects.NET.PictureProcessor")
  if err.number = 0 then
    Set objPictureProcessor = nothing
    Response.Write "FOUND: Smart Image Processor Own component<br>"
    ImageComponent = "PICPROC"	  
  else
    Response.Write "NOT FOUND: Smart Image Processor Own component<br>"  
  end if
	'Check for AspJpeg
  err.clear
  Set objASPjpeg = Server.CreateObject("Persits.Jpeg")
  if err.number = 0 then
    Set objASPjpeg = nothing
    Response.Write "FOUND: ASPJpeg Server Component<br>"
    ImageComponent = "ASPJPEG"
  else
    Response.Write "NOT FOUND: ASPJpeg Server Component<br>"
	end if
  'Check for AspImage
  err.clear
  Set AspImage = Server.CreateObject("AspImage.Image")
  if err.number = 0 then
    Set AspImage = nothing
    Response.Write "FOUND: ASPImage Server Component<br>"
    ImageComponent = "ASPIMAGE"
  else
    Response.Write "NOT FOUND: ASPImage Server Component<br>"
  end if
  'Check for AspSmart
	err.clear
	Set AspSmart = Server.CreateObject("aspSmartImage.SmartImage")
	if err.number = 0 then
	  set AspSmartImage = nothing
    Response.Write "FOUND: ASPSmartImage Server Component<br>"
	  ImageComponent = "ASPSMART"
	else
    Response.Write "NOT FOUND: ASPSmartImage Server Component<br>"
  end if
  'Check for ImgWriter
  err.clear
  Set objImgWriter = Server.CreateObject("softartisans.ImageGen")
  if err.number = 0 then
    Set objImgWriter = nothing
    Response.Write "FOUND: ImgWriter Server Component<br>"
    ImageComponent = "IMGWRITER"
  else
    Response.Write "NOT FOUND: ImgWriter Server Component<br>"
  end if
  'Check for AspThumb
  err.clear
  Set objAspThumb = Server.CreateObject("briz.AspThumb")
  if err.number = 0 then
    Set objAspThumb = nothing
    Response.Write "FOUND: AspThumb Server Component<br>"
    ImageComponent = "ASPTHUMB"
  else
    Response.Write "NOT FOUND: AspThumb Server Component<br>"
  end if
  on error goto 0
  
  NetImageComponent = DetectDotNetComponent(DotNetResize)
  if NetImageComponent <> "" then ImageComponent = NetImageComponent  
  
  DetectImageComponent = ImageComponent
end function

function DetectDotNetComponent(DotNetResize)
  Dim objHttp, DotNetImageComponent, ResizeComUrl, LastPath
    DotNetImageComponent = ""
    ResizeComUrl = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO")
    LastPath = InStrRev(ResizeComUrl,"/")
    if LastPath > 0 then
      ResizeComUrl = left(ResizeComUrl,Lastpath)
    end if
    ResizeComUrl = ResizeComUrl & DotNetResize
    
    'Check for ASP.NET 1
    on error resume next
    err.clear
    Set DotNet = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
    if err.number = 0 then
      objHttp.open "GET", ResizeComUrl, false
      objHttp.Send ""
      if trim(objHttp.responseText) <> "" and instr(objHttp.responseText,"@ Page Language=""C#""") = 0 then
        Response.Write "FOUND: ASP.NET Server Component<br>"
        DotNetImageComponent = "DOTNET1"
      end if
      Set DotNet = nothing
    else
      'Check for ASP.NET 2
      err.clear
      Set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
      if err.number = 0 then
        on error goto 0
        objHttp.open "GET", ResizeComUrl, false
        objHttp.Send ""
        if trim(objHttp.responseText) <> "" and trim(objHttp.responseText) = "DONE" then
          Response.Write "FOUND: ASP.NET Server Component<br>"
          DotNetImageComponent = "DOTNET2"
        end if
        Set objHttp = nothing
      else
        'Check for ASP.NET 3
        err.clear
        Set objHttp = Server.CreateObject("Microsoft.XMLHTTP")
        if err.number = 0 then
          objHttp.open "GET", ResizeComUrl, false
          objHttp.Send ""
          if trim(objHttp.responseText) <> "" and trim(objHttp.responseText) = "DONE" then
            Response.Write "FOUND: ASP.NET Server Component<br>"		  
            DotNetImageComponent = "DOTNET3"
          end if
          Set objHttp = nothing
        end if
      end if
    end if
    on error goto 0
  DetectDotNetComponent = DotNetImageComponent
end function

%>
<body>
<h1>Smart Image Processor Components</h1>
Detecting Components:<br><br>
<% det = DetectImageComponent("checkdotnet.aspx") %>
<h2>
<% if det <> "" then %>
  <a href="http://www.dmxzone.com/go?3965">Smart Image Processor</a> can be fully 
  used on this server!<br>
<% else %>
  Sorry but the <a href="http://www.dmxzone.com/go?3965">Smart Image Processor</a> 
  can not be used on this server yet.<br>
  You should install at least one supported server component. Preferably the included 
  own component or ASP.NET. <a href="http://www.dmxzone.com/go?3984">Read more 
  &gt;&gt;</a> 
  <% end if %>
</h2>
</body>
</html>
