<%@ Language=VBScript %>
<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %>
<% Response.Expires = -1 %>
<% Response.Buffer = True%>
<!--#include file="include/Connect.asp"--> <%' PDI control Constants%>
<!--#include file="include/Query.asp"--> <%' PDI Query Routines%>
<!--#include file="include/Header.inc"--> <%' miscellaneous server-side procedures%>
<!--#include file="include/Footer.inc"--> <%' miscellaneous server-side procedures%>
<!--#include file="include/UtilityProcs.asp"--> <%' miscellaneous server-side procedures%>
<%


    Session("SecurityOp") = "adm"
    if Session("LoginLevel")  < SecurityLevelNeeded(Session("SecurityOp")) then
        Response.Redirect "Login.asp"
	end if
%>

<html>
<head>
<link rel="stylesheet" href="pdi.css" type="text/css" />
<LINK rel="stylesheet" type="text/css" href="Stylesheets/Applicant.css">
<title>PDI/Applicant</title>
</head>
<!--#include file="include/Body.inc"--> 
</body>
</html>
<TABLE width=100%>
<tr>
	<td><B>MENU</B></td> <td align = "right"><A HREF="default.asp">Applicant Home Page</a></td>
</tr>
<tr>
	<td align="left"><B>&nbsp;&nbsp;<a href="WebPref.asp">Web Appearance and Application Preferences</a></td>
</tr>
<tr>
	<td align="left"><B>&nbsp;&nbsp;<a href="EditPage.asp">Web Page Text</a></td>
</tr>
<tr>
	<td align="left"><B>&nbsp;&nbsp;<a href="JobSetup.asp">Job Position Requirements</a></td>
</tr>
<tr>
	<td align="left"><B>&nbsp;&nbsp;<a href="UserQuestions.asp">User Defined Application Questions</a></td>
</tr>
<tr>
	<td align="left"><B>&nbsp;&nbsp;<a href="AdditionalInfo.asp">User Defined Applicant Information</a></td>
</tr>
<tr>
	<td align="left"><B>&nbsp;&nbsp;<a href="ScreenOptions.asp">Screening Options</a></td>
</tr>
<tr>
	<td align="left"><B>&nbsp;&nbsp;<a href="Forms.aspx">Forms</a></td>
</tr>


<%GetFooter%>