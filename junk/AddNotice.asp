<%@ Language=VBScript %>
<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %>
<% Response.Expires = -1 %>
<%Response.Buffer = True%>
<html>
<head>
<LINK rel="stylesheet" type="text/css" href="Stylesheets/Applicant.css">
<!--#include file="include/Connect.asp"--> <%' PDI control Constants%>
<!--#include file="include/Query.asp"--> <%' PDI control Constants%>
<!--#include file="include/UtilityProcs.asp"--> <%' miscellaneous server-side procedures%>
<!--#include file="include/Header.inc"--> <%' miscellaneous server-side procedures%>
<!--#include file="include/Footer.inc"--> <%' miscellaneous server-side procedures%>
<script language=javascript src="include/ClientProcs.js"></script>
<title>PDI/Applicant</title>

<%
' ******************************************************************************************************************************************
' ******************************************************************************************************************************************

	Dim Description
	Dim AllSites
	Dim Action	
	Dim SiteKey
	


	'** Load on initial entry or reload from an undo
	if "" & Request.Form("Action") = "" then
		if "" & Request("Operation") = "" then 
			response.redirect "Configuration.asp"
		end if
	end if

	QueryKey = 0

	SiteKey = clng("" & Request.Cookies("SiteDefaults")("DefaultSite"))

	if "" & Request.Form("Action") = "SAVE" or "" & Request.Form("Action") = "DELETE" then
		if "" & Request.Form("AllSites") = "" then
			QuerySiteKey = SiteKey
		else
			QuerySiteKey = 0
		end if
		if "" & Request.Form("QueryKey") <> ""  then  
			QueryKey =  Request.Form("QueryKey")
		else
			QueryKey = - 1
		end if
		with oCmd 
			Set .ActiveConnection = oConn
			.CommandText = "APP_ApplicantQuerySave_SP"
			.CommandType = adCmdStoredProc
			.Parameters.Append .CreateParameter("@SessionKey", adInteger,  adParamInput, 4, NULL)
			.Parameters.Append .CreateParameter("@ValidateOnly", adBoolean,  adParamInput, 1, 0)
			.Parameters.Append .CreateParameter("@ApplicantQueryKey", adDecimal,  adParamInputOutput, 9, QueryKey)
			.Parameters("@ApplicantQueryKey").Precision = 15
			.Parameters("@ApplicantQueryKey").NumericScale = 0
			
			.Parameters.Append .CreateParameter("@ApplicantQuerySiteKey", adDecimal,  adParamInputOutput, 9, QuerySiteKey) 
			.Parameters("@ApplicantQuerySiteKey").Precision = 15
			.Parameters("@ApplicantQuerySiteKey").NumericScale = 0
				
			.Parameters.Append .CreateParameter("@ApplicantQueryGUID", adbinary,  adParamInputOutput, 16, NULL)
			.Parameters.Append .CreateParameter("@ApplicantQueryDescription", adVarChar,  adParamInputOutput, 255, "" & Request.Form("Description"))
			.Parameters.Append .CreateParameter("@ApplicantQueryZip1", adVarChar,  adParamInputOutput, 10, "" & Request.Form("Zip1"))
			.Parameters.Append .CreateParameter("@ApplicantQueryZip2", adVarChar,  adParamInputOutput, 10, "" & Request.Form("Zip2"))
			.Parameters.Append .CreateParameter("@ApplicantQueryZip3", adVarChar,  adParamInputOutput, 10, "" & Request.Form("Zip3"))

			if "" & Request.Form("Position1") <> "" then
				.Parameters.Append .CreateParameter("@ApplicantQueryPosition1", adDecimal,  adParamInputOutput, 9, Request.Form("Position1"))
			else
				.Parameters.Append .CreateParameter("@ApplicantQueryPosition1", adDecimal,  adParamInputOutput, 9, NULL)
			end if
			.Parameters("@ApplicantQueryPosition1").Precision = 15
			.Parameters("@ApplicantQueryPosition1").NumericScale = 0

			if "" & Request.Form("Position2") <> "" then
				.Parameters.Append .CreateParameter("@ApplicantQueryPosition2", adDecimal,  adParamInputOutput, 9, Request.Form("Position2"))
			else
				.Parameters.Append .CreateParameter("@ApplicantQueryPosition2", adDecimal,  adParamInputOutput, 9, NULL)			
			end if
			.Parameters("@ApplicantQueryPosition2").Precision = 15
			.Parameters("@ApplicantQueryPosition2").NumericScale = 0

			if "" & Request.Form("Position3") <> "" then
				.Parameters.Append .CreateParameter("@ApplicantQueryPosition3", adDecimal,  adParamInputOutput, 9,  Request.Form("Position3"))
			else
				.Parameters.Append .CreateParameter("@ApplicantQueryPosition3", adDecimal,  adParamInputOutput, 9,  NULL)
			end if
			.Parameters("@ApplicantQueryPosition3").Precision = 15
			.Parameters("@ApplicantQueryPosition3").NumericScale = 0

			if "" & Request.Form("AppliedNumber") = "" then
				.Parameters.Append .CreateParameter("@ApplicantQueryAppliedNumber", adInteger,  adParamInputOutput, 4, NULL)
			else
				.Parameters.Append .CreateParameter("@ApplicantQueryAppliedNumber", adInteger,  adParamInputOutput, 4, Request.Form("AppliedNumber"))
			end if
			.Parameters.Append .CreateParameter("@ApplicantQueryAppliedTimeFrame", adTinyInt,  adParamInputOutput, 1, Request.Form("AppliedTimeFrame"))
			.Parameters.Append .CreateParameter("@ApplicantQueryAppliedAt", adTinyInt,  adParamInputOutput, 1,  Request.Form("AppliedAt"))
			if ShowScreen then
				.Parameters.Append .CreateParameter("@ApplicantQueryScreenStatus", adTinyInt,  adParamInputOutput, 1, Request.Form("ScreenStatus"))
			else
				.Parameters.Append .CreateParameter("@ApplicantQueryScreenStatus", adTinyInt,  adParamInputOutput, 1, 0)
			end if
			if "" & Request.Form("ScreenNumber") = "" then
				.Parameters.Append .CreateParameter("@ApplicantQueryScreenNumber", adInteger,  adParamInputOutput, 4, NULL)
			else

				.Parameters.Append .CreateParameter("@ApplicantQueryScreenNumber", adInteger,  adParamInputOutput, 4, Request.Form("ScreenNumber"))
			end if

			if "" & Request.Form("ScreenTimeFrame") <> "" then
				.Parameters.Append .CreateParameter("@ApplicantQueryScreenTimeFrame", adTinyInt,  adParamInputOutput, 1, Request.Form("ScreenTimeFrame"))
			else
				.Parameters.Append .CreateParameter("@ApplicantQueryScreenTimeFrame", adTinyInt,  adParamInputOutput, 1, NULL)
			end if

			if "" & Request.Form("ScoreLow") = "" then
				.Parameters.Append .CreateParameter("@ApplicantQueryScoreLow", adTinyInt,  adParamInputOutput, 1, NULL)
			else
				.Parameters.Append .CreateParameter("@ApplicantQueryScoreLow", adTinyInt,  adParamInputOutput, 1, Request.Form("ScoreLow"))
			end if
			if "" & Request.Form("ScoreHigh") = "" then
				.Parameters.Append .CreateParameter("@ApplicantQueryScoreHigh", adTinyInt,  adParamInputOutput, 1, NULL)
			else
				.Parameters.Append .CreateParameter("@ApplicantQueryScoreHigh", adTinyInt,  adParamInputOutput, 1,  Request.Form("ScoreHigh"))
			end if


			if "" & Request.Form("MathScoreLow") = "" then
				.Parameters.Append .CreateParameter("@ApplicantQueryMathScoreLow", adTinyInt,  adParamInputOutput, 1, NULL)
			else
				.Parameters.Append .CreateParameter("@ApplicantQueryMathScoreLow", adTinyInt,  adParamInputOutput, 1, Request.Form("MathScoreLow"))
			end if
			if "" & Request.Form("MathScoreHigh") = "" then
				.Parameters.Append .CreateParameter("@ApplicantQueryMathScoreHigh", adTinyInt,  adParamInputOutput, 1, NULL)
			else
				.Parameters.Append .CreateParameter("@ApplicantQueryMathScoreHigh", adTinyInt,  adParamInputOutput, 1,  Request.Form("MathScoreHigh"))
			end if

			If "" & Request.Form("Action") = "DELETE" then
				.Parameters.Append .CreateParameter("@Delete", adBoolean,  adParamInput, 1, True) 
			else
				.Parameters.Append .CreateParameter("@Delete", adBoolean,  adParamInput, 1, False) 
			End if

			if "" & Request.Form("MinimumAge") = "" then
				.Parameters.Append .CreateParameter("@MinimumAge", adTinyInt,  adParamInputOutput, 1, NULL)
			else
				.Parameters.Append .CreateParameter("@MinimumAge", adTinyInt,  adParamInputOutput, 1, Request.Form("MinimumAge"))
			end if

			if "" & Request.Form("UpdatedNumber") = "" then
				.Parameters.Append .CreateParameter("@ApplicantQueryUpdatedNumber", adInteger,  adParamInputOutput, 4, NULL)
			else
				.Parameters.Append .CreateParameter("@ApplicantQueryUpdatedNumber", adInteger,  adParamInputOutput, 4, Request.Form("UpdatedNumber"))
			end if
			.Parameters.Append .CreateParameter("@ApplicantQueryUpdatedTimeFrame", adTinyInt,  adParamInputOutput, 1, Request.Form("UpdatedTimeFrame"))


			.Execute ,, adExecuteNoRecords

		End With		 			
		QueryKey = clng(oCmd.Parameters("@ApplicantQueryKey").Value)
		Action = "SAVE"
	end if

	'** Load on initial entry or reload from an undo
	if "" & Request.Form("Action") = "UNDO" then
		QueryKey =  Request.Form("Key")
	else
		if "" & Request("Operation") = "EDIT" then 
			QueryKey =  Request("Key")
		end if
	end if

	if QueryKey <> 0 then
		oRs.Open "SELECT * from Applicant_Queries WHERE Applicant_Query_Key = " & QueryKey , oConn, adOpenForwardOnly, adLockReadOnly
		
		AllSites = false
		If Not oRs.EOF Then
			LoadQueryValues(oRS)	
			if clng(QuerySiteKey) = 0 then AllSites = True
		end if
		oRS.Close
	end if

			
%>
<script LANGUAGE="javascript">
<!--
//**********************************************************************************************************************
//**********************************************************************************************************************

	function EnableDate()
	{
		var oForm = document.forms[0];
		if (oForm.elements["ScreenStatus"].value==7||oForm.elements["ScreenStatus"].value==4||oForm.elements["ScreenStatus"].value==8)
		{
			oForm.elements["ScreenNumber"].disabled=false;
			oForm.elements["ScreenNumber"].style.backgroundColor = "#99ccff";
			oForm.elements["ScreenTimeFrame"].disabled=false;
			oForm.elements["ScreenTimeFrame"].style.backgroundColor = "#99ccff";
		}
		else
		{
			oForm.elements["ScreenNumber"].disabled=true;
			oForm.elements["ScreenNumber"].style.backgroundColor = "LightGrey";
			oForm.elements["ScreenTimeFrame"].disabled=true;
			oForm.elements["ScreenTimeFrame"].style.backgroundColor = "LightGrey";
		}
		
	}

	function Undo()
	{
		var oForm = document.forms[0];
		oForm.elements["Action"].value = "UNDO";
		oForm.submit();
	}

	function Save()
	{
		var oForm = document.forms[0];
		if (oForm.elements["Description"].length==0||oForm.elements["Description"].value=="")
		{
			alert("You must provide a description before a notice can be saved.");
			return true;
		}
		oForm.elements["Action"].value = "SAVE";
		oForm.submit();
	}

	function Delete()
	{
		var oForm = document.forms[0];
		oForm.elements["Action"].value = "DELETE";
		oForm.submit();
	}

//**********************************************************************************************************************
//**********************************************************************************************************************
//-->
</script>

</head>
<!--#include file="include/Body.inc"--> 
<TABLE>
<tr>
<td align="left">

<FORM id="frmConfiguration" method="post" action="<%=Request.ServerVariables("PATH_INFO")%>">
<input type="hidden" id="Action" name="Action" value="<%=Action%>">
<script LANGUAGE="javascript">
<!--
	var oForm = document.forms[0];
	if (oForm.elements["Action"].value == "SAVE"||oForm.elements["Action"].value == "DELETE")
	{
		window.opener.document.forms[0].submit();
		window.close();
	}
//-->
</script>
<input type="hidden" id="QueryKey" name="QueryKey" value="<%=QueryKey %>">
<TABLE cols ="3">
  	<tr>
  		<td nowrap valign="center" align="right">Notice Description</td>
  		<td nowrap  align="left" colspan="2"><input type="text" AUTOCOMPLETE="OFF" class="InputText" id="Description" name="Description" value="<%=Description%>" size="74"</td>
 	</tr>

	<tr>
		<td nowrap  align="right">Zip Code(s)</td>
		<td nowrap  align="left"  colspan="2">
			<input type="text" AUTOCOMPLETE="OFF" class="InputText" id="Zip1" name="Zip1" size="7" MAXLENGTH="5" value="<%=Zip1%>" onkeypress="javascript:return ForceZipChars();">&nbsp;&nbsp;&nbsp
			<input type="text" AUTOCOMPLETE="OFF" class="InputText" id="Zip2" name="Zip2" size="7" MAXLENGTH="5" value="<%=Zip2%>" onkeypress="javascript:return ForceZipChars();">&nbsp;&nbsp;&nbsp
			<input type="text" AUTOCOMPLETE="OFF" class="InputText" id="Zip3" name="Zip3" size="7" MAXLENGTH="5" value="<%=Zip3%>" onkeypress="javascript:return ForceZipChars();">&nbsp;&nbsp;&nbsp
		</td>
	</tr>
	<tr>
  	<td nowrap  align="right">Position(s)</td>
      	<td nowrap  align="left" colspan="2"><select id="Position1" name="Position1" class="select150">
		<option value></option>		
  <%
  
  	oRs.Open "SELECT JobPosition_Key, JobPosition_Description FROM JobPositions ORDER BY JobPosition_Description ASC", oConn, adOpenDynamic, adLockReadOnly

	HasRec = false
	Do Until oRs.EOF
		HasRec = True
		If cStr("" & Position1) = cStr(oRs.Fields("JobPosition_Key").Value) Then 
		%>
		<option value="<%=oRs.Fields("JobPosition_Key").Value%>" SELECTED><%=oRs.Fields("JobPosition_Description").Value%></option>
		<%Else%>
		<option value="<%=oRs.Fields("JobPosition_Key").Value%>"><%=oRs.Fields("JobPosition_Description").Value%></option>		
		<%
		End If
		oRs.MoveNext
	Loop
	
  %>
      </select>	

      <select id="Position2" name="Position2" class="select150">
		<option value></option>		
  <%

	if HasRec then oRS.Movefirst   
	Do Until oRs.EOF
		
		If cStr("" & Position2) = cStr(oRs.Fields("JobPosition_Key").Value) Then 
		%>
		<option value="<%=oRs.Fields("JobPosition_Key").Value%>" SELECTED><%=oRs.Fields("JobPosition_Description").Value%></option>
		<%Else%>
		<option value="<%=oRs.Fields("JobPosition_Key").Value%>"><%=oRs.Fields("JobPosition_Description").Value%></option>		
		<%
		End If
		oRs.MoveNext
	Loop
	
  %>
	</select>	
      <select id="Position3" name="Position3" class="select150">
		<option value></option>		
  <%
  
	if HasRec then oRS.Movefirst   
	Do Until oRs.EOF

		If cStr("" & Position3) = cStr(oRs.Fields("JobPosition_Key").Value) Then 
		%>
			<option value="<%=oRs.Fields("JobPosition_Key").Value%>" SELECTED><%=oRs.Fields("JobPosition_Description").Value%></option>
		<%Else%>
			<option value="<%=oRs.Fields("JobPosition_Key").Value%>"><%=oRs.Fields("JobPosition_Description").Value%></option>		
		<%
		End If
		oRs.MoveNext
	Loop

	oRs.Close
  %>
	</select>	
</td>
 </tr>
 <tr>
	<td nowrap  align="right" valign="center">Applied within the last</td>
	<td nowrap  align="left" valign="top"><input type="text"  AUTOCOMPLETE="OFF" class="InputText"  MAXLENGTH="3" id="AppliedNumber" name="AppliedNumber" onkeypress="javascript:return ForceDigitChars();" size="3" value="<%=AppliedNumber%>">
        <select id="AppliedTimeframe" name="AppliedTimeframe" class="select100">
  <%

    	oRs.Open "SELECT 1 as SelKey, 'Days' SelDesc union select 2, 'Weeks' union select 3, 'Months' union select 4, 'Years'", oConn, adOpenForwardOnly, adLockReadOnly

	Do Until oRs.EOF 
		If cStr(AppliedTimeFrame) = cStr(oRS("SelKey")) Then %>
			<option value="<%=oRs.Fields("SelKey").Value%>" SELECTED><%=oRs.Fields("SelDesc").Value%></option>		
		<%else%>
			<option value="<%=oRs.Fields("SelKey").Value%>"><%=oRs.Fields("SelDesc").Value%></option>		
		<%end if
		oRs.MoveNext
	Loop
	
	oRs.Close
  %>
	</select>
	</td>	
 </tr>

<tr class="Collapse" nowrap>
	<td nowrap align="right" valign="center">Updated within the last</td>
	<td nowrap align="left" valign="top"><input type="text"  AUTOCOMPLETE="OFF" class="InputText"   MAXLENGTH="3" id="Text1" name="UpdatedNumber" onkeypress="javascript:return ForceDigitChars();" size="3" value="<%=UpdatedNumber%>">
        <select id="UpdatedTimeFrame" name="UpdatedTimeFrame" class="select100" value="<%=UpdatedTimeFrame%>">
  <%
   	oRs.Open "SELECT 1 as SelKey, 'Days' SelDesc union select 2, 'Weeks' union select 3, 'Months' union select 4, 'Years'", oConn, adOpenForwardOnly, adLockReadOnly
	Do Until oRs.EOF 
		If cStr("" & UpdatedTimeFrame) = cStr(oRS("SelKey")) Then %>
			<option value="<%=oRs.Fields("SelKey").Value%>" SELECTED><%=oRs.Fields("SelDesc").Value%></option>		
		<%else%>
			<option value="<%=oRs.Fields("SelKey").Value%>"><%=oRs.Fields("SelDesc").Value%></option>		
		<%end if
		oRs.MoveNext
	Loop
	oRs.Close
  %>
	</select>
	</td>	
    </tr>

<tr>
	<td nowrap  align="right" valign="center">Applied at</td>
        <td nowrap  align="left" colspan="2"><select id="AppliedAt" name="AppliedAt" class="select100">
  <%
 
  	oRs.Open "SELECT 0 as SelKey, '(Any Site)' SelDesc union select 1, '(This Site)'", oConn, adOpenForwardOnly, adLockReadOnly

	Do Until oRs.EOF 
		If cStr(AppliedAt) = cStr(oRS("SelKey")) Then 
		%>
			<option value="<%=oRs.Fields("SelKey").Value%>" SELECTED><%=oRs.Fields("SelDesc").Value%></option>	
		<%else%>
			<option value="<%=oRs.Fields("SelKey").Value%>"><%=oRs.Fields("SelDesc").Value%></option>		
		<%end if
		oRs.MoveNext
	Loop
	
	oRs.Close
  %>
	</select>	


<%if ShowScreen then%>	
<tr>
	<td nowrap  valign="center" align="right">Background Screen Status</td>
	<td nowrap  valign="bottom" colspan="2">
		<select id="ScreenStatus" name="ScreenStatus" class="select150" onchange="javascript:return EnableDate();">		
		<%
		Dim SQL
		SQL = "Select 0 as Status_Key, '(Any Status)' as Status_Desc"
		SQL = SQL & " Union Select 1 as Status_Key, 'No Screen' as Status_Desc"
		SQL = SQL & " Union Select 2 as Status_Key, 'Requested' as Status_Desc"
		SQL = SQL & " Union Select 3 as Status_Key, 'Sent' as Status_Desc"
		SQL = SQL & " Union Select 4 as Status_Key, 'Cleared' as Status_Desc"
		SQL = SQL & " Union Select 6 as Status_Key, 'Under Review' as Status_Desc"
		SQL = SQL & " Union Select 7 as Status_Key, 'Do Not Hire' as Status_Desc"
		SQL = SQL & " Union Select 8 as Status_Key, 'Cleared or Do Not Hire' as Status_Desc"
		oRs.Open SQL,  oConn, adOpenForwardOnly, adLockReadOnly
		Do Until oRs.EOF
			if cstr(ScreenStatus) = "" then ScreenStatus = 0
			If cstr("" &  ScreenStatus) = cstr(oRs.Fields("Status_Key").Value) Then
			%>
				<option style="width:125px;" value="<%=oRs.Fields("Status_Key").Value%>" SELECTED><%=oRs.Fields("Status_Desc").Value%></option>		
			<%Else%>
				<option style="width:125px;" value="<%=oRs.Fields("Status_Key").Value%>"><%=oRs.Fields("Status_Desc").Value%></option>		
			<%end if
			oRs.MoveNext
		Loop
		oRs.Close
	  	%>
		</select>
	</td>
</tr>
<tr>
	<td nowrap  align="right" valign="center">Screen Completed Within the Last</td>
	<td nowrap  align="left" valign="top" colspan="2"><input type="text"  AUTOCOMPLETE="OFF" MAXLENGTH="3" class="InputText"  id="ScreenNumber" name="ScreenNumber" onkeypress="javascript:return ForceDigitChars();" size="3" value="<%=ScreenNumber%>">
        <select id="ScreenTimeFrame" name="ScreenTimeFrame" class="select100">
  <%
  
  	oRs.Open "SELECT 1 as SelKey, 'Days' SelDesc union select 2, 'Weeks' union select 3, 'Months' union select 4, 'Years'", oConn, adOpenForwardOnly, adLockReadOnly

	Do Until oRs.EOF 
		If cstr("" &  ScreenTimeFrame) = cstr(oRs.Fields("SelKey").Value) Then%>
			<option value="<%=oRs.Fields("SelKey").Value%>" SELECTED><%=oRs.Fields("SelDesc").Value%></option>	
		<%else%>
			<option value="<%=oRs.Fields("SelKey").Value%>"><%=oRs.Fields("SelDesc").Value%></option>		
		<%end if
		oRs.MoveNext
	Loop
	
	oRs.Close
  %>
	</select>	
  </td>
 </tr>
<%end if%>
	 <tr>
	  <td nowrap  align="right">PASS III scores from </td>
	  <td nowrap  colspan="4"><input type="text" AUTOCOMPLETE="OFF" class="InputText" id="ScoreLow" name="ScoreLow" MAXLENGTH="3" onkeypress="javascript:return ForceDigitChars();" size="3" value="<%=ScoreLow%>"> to <input type="text" AUTOCOMPLETE="OFF" class="InputText" id="ScoreHigh" name="ScoreHigh" MAXLENGTH="3" onkeypress="javascript:return ForceDigitChars();" size="3" value="<%=ScoreHigh%>"></td>
	 </tr>
	 <tr>
	  <td nowrap align="right">Math test scores from </td>
	  <td nowrap colspan="4"><input type="text" AUTOCOMPLETE="OFF" class="InputText" id="MathScoreLow" name="MathScoreLow"  MAXLENGTH="3" onkeypress="javascript:return ForceDigitChars();" size="3" value="<%=MathScoreLow%>"> to <input type="text" AUTOCOMPLETE="OFF" class="InputText" id="MathScoreHigh" name="MathScoreHigh"  MAXLENGTH="3" onkeypress="javascript:return ForceDigitChars();" size="3" value="<%=MathScoreHigh%>"></td>
	 </tr>

  <% oRs.Open "SELECT WebPref_ShowDOB FROM WebPrefs", oConn, adOpenForwardOnly, adLockReadOnly
	 Dim ShowMinimumAge
	 ShowMinimumAge = oRs.Fields("WebPref_ShowDOB").Value
     oRs.Close
     if ShowMinimumAge = 1 or ShowMinimumAge then %>		

	<tr class="Collapse" nowrap>
	  <td nowrap align="right">Minimum Age</td>
	  <td nowrap colspan="4">
		<select id="MinimumAge" name="MinimumAge" class="select75">		
  <%
 
  	oRs.Open "SELECT 0 AS SelKey, '(Any Age)' SelDesc UNION SELECT 15, '15' UNION SELECT 16, '16' UNION SELECT 17, '17' UNION SELECT 18, '18' UNION SELECT 19, '19' UNION SELECT 20, '20' UNION SELECT 21, '21' ORDER BY SelDesc", oConn, adOpenForwardOnly, adLockReadOnly

	Do Until oRs.EOF 
	if ((cStr("" & MinimumAge) & "" = oRs.Fields("SelKey").Value & "")) then%>
			<option value="<%=oRs.Fields("SelKey").Value%>" selected="selected"><%=oRs.Fields("SelDesc").Value%></option>		
	<%else%>
			<option value="<%=oRs.Fields("SelKey").Value%>"><%=oRs.Fields("SelDesc").Value%></option>		
	<%end if%>
	<%
			oRs.MoveNext
	Loop
	
	oRs.Close
  end if
  %>
			
	</select>
	 </td>
	 </tr>

	 <tr>
	  <td align="right"></td>
	<%if AllSites then %>
	   <td nowrap  colspan="4"><input type="checkbox" AUTOCOMPLETE="OFF" id="AllSites" name="AllSites" value="<%=AllSites%>" CHECKED>This notice is available at all sites.</td>
	<%else%>
	   <td nowrap  colspan="4"><input type="checkbox" AUTOCOMPLETE="OFF" id="AllSites" name="AllSites" value="<%=AllSites%>">This notice is available at all sites.</td>
	<%end if%>
	 </tr>

</tr>
 <tr>
  <td colspan="3" align="center"><hr></td>
 </tr>
<tr>
  <td nowrap  colspan="3" align="center">
	<input type="button" class="navbutton" style="height:30px" value="Save" id="btnSave" name="btnSave" onclick="javascript:Save();">&nbsp;&nbsp;
	<input type="button" class="navbutton" style="height:30px" value="Undo" id="btnUndo" name="btnUndo" onclick="javascript:Undo();">&nbsp;&nbsp;
	<input type="button" class="navbutton" style="height:30px" value="Delete" id="btnDelete" name="btnDelete" onclick="javascript:Delete();">&nbsp;&nbsp;
	<input type="button" class="navbutton" style="height:30px" value="Quit" id="btnQuit" name="btnQuit" onclick="javascript:window.close();">
	</td>
</tr>

</TABLE>
</FORM>
</table>
</body>
<SCRIPT LANGUAGE="vbscript">
	'****************************************************************************************************
	'****************************************************************************************************
	
	Sub SetLostFocusColor()
		
		Dim oCurrentCtrl
		
		Set oCurrentCtrl = window.event.srcElement
		
		oCurrentCtrl.style.backgroundColor = "#99ccff"
		
	End Sub
	
	'****************************************************************************************************
	'****************************************************************************************************
	
	Sub SetGotFocusColor()

		Dim oCurrentCtrl
		
		Set oCurrentCtrl = window.event.srcElement
		
		oCurrentCtrl.style.backgroundColor = "#00ffff"

	End Sub	

	'****************************************************************************************************
	'****************************************************************************************************
		
	Dim x
	Dim oGotFocus
	Dim oLostFocus
	Dim oForm
		
	x = 0
	
	Set oForm = document.forms(0)
	Set oGotFocus = getRef("SetGotFocusColor")
	Set oLostFocus = getRef("SetLostFocusColor")
	
	For x = 0 to oForm.length - 1
		Select Case oForm.item(x).type
			Case "text", "select-one"
				oForm.item(x).attachEvent "onfocus", oGotFocus
				oForm.item(x).attachEvent "onblur", oLostFocus			
			Case Else ' do nothing
			
		End Select

	Next
</SCRIPT>
<%
	' Clean-up
	
	On Error Resume Next

	oRs.Close
	Set oRs = Nothing
	
	oCmd.Cancel
	Set oCmd = Nothing
	
	oConn.Close
	Set oConn = Nothing
%>
<%if ShowScreen then%>
<script LANGUAGE="javascript">
<!--
	EnableDate();
//-->
</script>
<%end if %>

</html>
<!--#include file="include/FieldProps.inc"--> 
<SCRIPT LANGUAGE="vbscript">
	on error resume next
	LinkFocus()
	ResetDisplay()
	document.forms(0).item("Description").focus
</SCRIPT>