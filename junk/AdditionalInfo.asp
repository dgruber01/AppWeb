<%@ Language=VBScript %>
<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %>
<% Response.Expires = -1 %>
<% Response.Buffer = True%>
<!--#include file="include/Connect.asp"--> <%' PDI control Constants%>
<!--#include file="include/Query.asp"--> <%' PDI Query Routines%>
<!--#include file="include/Header.inc"--> <%' miscellaneous server-side procedures%>
<!--#include file="include/Footer.inc"--> <%' miscellaneous server-side procedures%>
<% 
	Dim PageKey 
	Dim Qno
	Dim oRSOrder
	Dim DefQ
	Dim DefDir

	Qno  = 1
	DefQ = 1
	DefDir = "DOWN"
	RecCount = 0

	if Request.form("Action") = "MOVE" then
		oRS.open "Select * from WebInfo Where isnull(WebInfo_Deleted, 0) <> 1 order by WebInfo_Order" , oConn, adOpenDynamic, adLockOptimistic
		Do while not oRS.eof 
			RecCount = RecCount + 1
			oRS.movenext
		loop
		if RecCount > 0 then oRs.movefirst
		For X = 1 to clng(Request.form("Qno")) - 1
			if oRS.eof then break
			oRS.movenext
		next
		if not oRS.eof then
			Set oRSOrder = Server.CreateObject("ADODB.Recordset")
			if Request.form("dir") = "UP" then
				oRSOrder.open "Select * from WebInfo where isnull(WebInfo_Deleted, 0) <> 1  and  WebInfo_Order <" & oRS("WebInfo_Order") & " order by  WebInfo_Order DESC",  oConn, adOpenDynamic, adLockOptimistic
				DefQ = clng(Request.form("Qno")) - 1
				if DefQ > 1 then
					DefDir = "UP"
				else
					DefDir = "DOWN"
				end if
			else
				oRSOrder.open "Select * from WebInfo where isnull(WebInfo_Deleted, 0) <> 1  and WebInfo_Order >" & oRS("WebInfo_Order") & " order by  WebInfo_Order",  oConn, adOpenDynamic, adLockOptimistic
				DefQ = clng(Request.form("Qno")) + 1
				if DefQ < RecCount then
					DefDir = "DOWN"
				else
					DefDir = "UP"

				end if

			end if
			if not oRSOrder.eof then
				SaveOrder = oRSOrder("WebInfo_Order")
				oRSOrder("WebInfo_Order") = oRS("WebInfo_Order")
				oRS("WebInfo_Order") = SaveOrder
				oRSOrder.update
				oRS.update
			end if
			oRSOrder.close
			Set oRSOrder = nothing
		end if
		oRS.close
	end if

%>
<script language=javascript>

	function SelectPosition(ctl)
	{
		var oForm = document.forms[0];
		oForm.submit();
	}

	function movequestion()
	{
		var oForm = document.forms[0];
		oForm.elements("Action").value = "MOVE";
		oForm.submit();

	}
	

</script>

<html>
<head>
<link rel="stylesheet" href="pdi.css" type="text/css" />
<LINK rel="stylesheet" type="text/css" href="Stylesheets/Applicant.css">
<title>PDI/Applicant</title>
</head>
<!--#include file="include/Body.inc"--> 
<input type="hidden" id="Action" name="Action">
<table width = "100%">
<tr>
<td></td><td align="right"><a HREF="default.asp">Applicant Home Page</a></td>
</tr>
<form method="post">
<input type="hidden" id="Action" name="Action">
<tr>
<td >&nbsp;
</td><td align="right"><a HREF="AppAdmin.asp">Applicant Web Preferences Page</a></td></table>
	<%
	oRS.open "Select * from WebInfo Where isnull(WebInfo_Deleted, 0) <> 1  order by WebInfo_Order " 
	do while not oRS.eof %>
		<%QuestionID = " (IQ" & oRS("WebInfo_Key") & ")"%>
		<%if len(trim("" & oRS("WebInfo_Question_Text"))) = 0 then %> 
			<p><%=Qno%>.&nbsp; <A href="EditInfo.asp?Key=<%=oRS("WebInfo_Key")%>">No Text Entered<%=QuestionID%></A></p>
		<%else%>
			<p><%=Qno%>.&nbsp;<A href="EditInfo.asp?Key=<%=oRS("WebInfo_Key")%>"><%=oRS("WebInfo_Question_Text")%><%=QuestionID%></A></p>
		<%end if%>
	<%	Qno = Qno + 1
		oRS.movenext
	loop
	oRS.close
	%>

	<p>Move Question
	<select id="qno" name="qno" >
	<%for X = 1 to Qno - 1%>
		<%if DefQ = X then%>
			<option value = "<%=X%>" SELECTED><%=X%>
		<%else%>
			<option value = "<%=X%>"><%=X%>
		<%end if%>
	</Option><%next%></select>
	<select id="dir" name="dir">
		<%if DefDir = "DOWN" then%>
			<option value = "Down" SELECTED>Down</Option>
			<option value = "UP">Up</Option>
		<%else%>
			<option value = "Down" >Down</Option>
			<option value = "UP" SELECTED>Up</Option>
		<%end if%>

	</select> <input type="button" class="navbutton" style="width:50px;height:20px" value="Move" onclick="javascript:movequestion();"</p>
	<P valign="top" colspan="8" align="center">
		<input class="navbutton" style="height=20px;width:80px" type="button" value="Add Field" name="AddQuestion" onclick="javascript:return window.navigate('EditInfo.asp');">
	</P>
</form>
</body>
</html>
<%GetFooter%>
<!--#include file="include/FieldProps.inc"--> 
<SCRIPT LANGUAGE="vbscript">
	on error resume next
	LinkFocus()
	ResetDisplay()
</SCRIPT>