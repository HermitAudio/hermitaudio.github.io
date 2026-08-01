<%@ LANGUAGE="VBSCRIPT" %>
<%
 Option Explicit
 Dim ranWizard, intID, i, background, theme, frameHeight, pagePicture, pageText, pageType, pagewords

 If myinfo.Theme <> "" Then
	Theme = myinfo.Theme
%>
<!--	$Date: 9/05/97 11:21a $	$ModTime: $	$Revision: 8 $	$Workfile: default.asp $-->
	<html>
	<head>
	<!--#include virtual ="/iissamples/homepage/sub.inc"-->
	<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
	<meta HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1">
	<title><% call Title %></title>
	<% call styleSheet %>
	<meta name="Microsoft Border" content="l, default"></head>
	<body TopMargin="0" Leftmargin="0"><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td valign="top" width="1%">

<p>&nbsp;</p>

<p align="center"><font face="Arial"><strong>ShortCuts</strong></font></p>

<p>
<applet code="fphover.class" codebase="_fpclass/" width="120" height="24">
  <param name="text" value="The Audio Site">
  <param name="color" value="#000080">
  <param name="hovercolor" value="#00FFFF">
  <param name="textcolor" value="#FFFFFF">
  <param name="effect" value="glow">
  <param name="url" value="AudioStartsHere.htm" valuetype="ref">
</applet>
<applet code="fphover.class" codebase="_fpclass/" width="120" height="24">
  <param name="text" value="Electrocompaniet">
  <param name="color" value="#000080">
  <param name="hovercolor" value="#00FFFF">
  <param name="textcolor" value="#FFFFFF">
  <param name="effect" value="glow">
  <param name="url" value="OtalaStory.htm" valuetype="ref">
</applet>
<applet code="fphover.class" codebase="_fpclass/" width="120" height="24">
  <param name="text" value="Audio Theory">
  <param name="color" value="#000080">
  <param name="hovercolor" value="#00FFFF">
  <param name="textcolor" value="#FFFFFF">
  <param name="effect" value="glow">
  <param name="url" value="theory_of_single_stages.htm" valuetype="ref">
</applet>
<applet code="fphover.class" codebase="_fpclass/" width="120" height="24">
  <param name="text" value="Schematics">
  <param name="color" value="#000080">
  <param name="hovercolor" value="#00FFFF">
  <param name="textcolor" value="#FFFFFF">
  <param name="effect" value="glow">
  <param name="url" value="schematics.htm" valuetype="ref">
</applet>
&nbsp;&nbsp; 
<applet code="fphover.class" codebase="_fpclass/" width="120" height="24">
  <param name="text" value="Osiris Data AS">
  <param name="color" value="#000080">
  <param name="hovercolor" value="#00FFFF">
  <param name="textcolor" value="#FFFFFF">
  <param name="effect" value="glow">
  <param name="url" value="http://www.osiris.no" valuetype="ref">
  <param name="font" value="Dialog">
  <param name="fontstyle" value="bold">
  <param name="fontsize" value="14">
</applet>
<applet code="fphover.class" codebase="_fpclass/" width="120" height="24">
  <param name="text" value="Content">
  <param name="color" value="#000080">
  <param name="hovercolor" value="#00FFFF">
  <param name="textcolor" value="#FFFFFF">
  <param name="effect" value="glow">
  <param name="url" value="TableOfContent.htm" valuetype="ref">
</applet>
<applet code="fphover.class" codebase="_fpclass/" width="120" height="24">
  <param name="text" value="News">
  <param name="color" value="#FF0000">
  <param name="hovercolor" value="#00FFFF">
  <param name="textcolor" value="#000000">
  <param name="effect" value="glow">
  <param name="url" value="news.htm" valuetype="ref">
  <param name="font" value="Dialog">
  <param name="fontstyle" value="regular">
  <param name="fontsize" value="14">
</applet>
</p>
</td><td valign="top" width="24"></td><!--msnavigation--><td valign="top">
	<!--#include virtual ="/iissamples/homepage/theme.inc"-->
	<!--msnavigation--></td></tr><!--msnavigation--></table></body>
	</html>
<%
 Else
	response.redirect "/IISSamples/Default/welcome.htm"
 End If
%>