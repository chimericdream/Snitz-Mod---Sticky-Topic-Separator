<%
'#################################################################################
'## Snitz Forums 2000 v3.4.06
'#################################################################################
'## Copyright (C) 2000-06 Michael Anderson, Pierre Gorissen,
'##                       Huw Reddick and Richard Kinser
'##
'## This program is free software; you can redistribute it and/or
'## modify it under the terms of the GNU General Public License
'## as published by the Free Software Foundation; either version 2
'## of the License, or (at your option) any later version.
'##
'## All copyright notices regarding Snitz Forums 2000
'## must remain intact in the scripts and in the outputted HTML
'## The "powered by" text/logo with a link back to
'## http://forum.snitz.com in the footer of the pages MUST
'## remain visible when the pages are viewed on the internet or intranet.
'##
'## This program is distributed in the hope that it will be useful,
'## but WITHOUT ANY WARRANTY; without even the implied warranty of
'## MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'## GNU General Public License for more details.
'##
'## You should have received a copy of the GNU General Public License
'## along with this program; if not, write to the Free Software
'## Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'##
'## Support can be obtained from our support forums at:
'## http://forum.snitz.com
'##
'## Correspondence and Marketing Questions can be sent to:
'## manderson@snitz.com
'##
'#################################################################################
%>
<!--#INCLUDE FILE="config.asp"-->
<%
Dim ArchiveView
Dim HeldFound, UnApprovedFound, UnModeratedPosts, UnModeratedFPosts
Dim HasHigherSub
HasHigherSub = false
'#################################################################################

if (Request("FORUM_ID") = "" or IsNumeric(Request("FORUM_ID")) = False) and (Request.Form("Method_Type") <> "login") and (Request.Form("Method_Type") <> "logout") then
	Response.Redirect "default.asp"
else
	Forum_ID = cLng(Request("FORUM_ID"))
end if 

'-------------------------------------------
' FORUM SORTING MOD VARIABLES
'-------------------------------------------

' Code Mod for mypage variable
dim mypage : mypage = request("whichpage")
if ((Trim(mypage) = "") or IsNumeric(mypage) = False) then mypage = 1
mypage = cLng(mypage)

' Topic Sorting Variables
dim strtopicsortord :strtopicsortord = request("sortorder")
dim strtopicsortfld :strtopicsortfld = request("sortfield")
dim strtopicsortday :strtopicsortday = request("days")
dim inttotaltopics : inttotaltopics = 0
dim strSortCol, strSortOrd

Select Case strtopicsortord
	Case "asc"
		strSortOrd = " ASC"
	Case Else
		strSortOrd = " DESC"
		strtopicsortord = "desc"
End Select

Select Case strtopicsortfld
	Case "topic"
		strSortCol = "T_SUBJECT" & strSortOrd
	Case "author"
		strSortCol = "M_NAME" & strSortOrd
	Case "replies"
		strSortCol = "T_REPLIES" & strSortOrd
	Case "views"
		strSortCol = "T_VIEW_COUNT" & strSortOrd
	Case "lastpost"
		strSortCol = "T_LAST_POST" & strSortOrd
	Case Else
		strtopicsortfld = "lastpost"
		strSortCol = "T_LAST_POST" & strSortOrd
End Select
strQStopicsort = "FORUM_ID=" & Forum_ID
'-------------------------------------------
if request("ARCHIVE") = "true" then
	strActivePrefix = strTablePrefix & "A_"
	ArchiveView = "true"
	ArchiveLink = "ARCHIVE=true&"
elseif request("ARCHIVE") <> "" then
	Response.Redirect "default.asp"
	Response.End
else
	strActivePrefix = strTablePrefix
	ArchiveView = ""
	ArchiveLink = ""
end if
%>
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_func_chknew.asp" -->
<!--#INCLUDE FILE="inc_subscription.asp" -->
<!--#INCLUDE FILE="inc_moderation.asp" -->
<%

nDays = Request.Cookies(strCookieURL & "NumDays")

if Request.form("cookie") = 1 then
	if strSetCookieToForum = "1" then
		Response.Cookies(strCookieURL & "NumDays").Path = strCookieURL
	end if
	Response.Cookies(strCookieURL & "NumDays") = Request.Form("days")
	Response.Cookies(strCookieURL & "NumDays").expires = dateAdd("yyyy", 1, strForumTimeAdjust)
	nDays = Request.Form("Days")
	mypage = 1
end if

if request("ARCHIVE") = "true" then
	nDays = "0"
end if

Response.Write	"    <script language=""JavaScript"" type=""text/javascript"">" & vbNewLine & _
		"    function ChangePage(fnum){" & vbNewLine & _
		"    	if (fnum == 1) {" & vbNewLine & _
		"    		document.PageNum1.submit();" & vbNewLine & _
		"    		}" & vbNewLine & _
		"    	else {" & vbNewLine & _
		"    		document.PageNum2.submit();" & vbNewLine & _
		"    	}" & vbNewLine & _
		"    }" & vbNewLine & _
		"    </script>" & vbNewLine


if mLev = 4 then
	AdminAllowed = 1
	ForumChkSkipAllowed = 1
elseif mLev = 3 then
	if chkForumModerator(Forum_ID, chkString(strDBNTUserName,"decode")) = "1" then
	 	AdminAllowed = 1
		ForumChkSkipAllowed = 1
	else   
		if lcase(strNoCookies) = "1" then
	 		AdminAllowed = 1
			ForumChkSkipAllowed = 0
		else
			AdminAllowed = 0
			ForumChkSkipAllowed = 0
		end if
	end if
elseif lcase(strNoCookies) = "1" then
	AdminAllowed = 1
	ForumChkSkipAllowed = 0
else
	AdminAllowed = 0
	ForumChkSkipAllowed = 0
end if

if strPrivateForums = "1" and (Request.Form("Method_Type") <> "login") and (Request.Form("Method_Type") <> "logout") and ForumChkSkipAllowed = 0 then
	result = ChkForumAccess(Forum_ID, MemberID, true)
end if

if strModeration = "1" and AdminAllowed = 1 then
	UnModeratedPosts = CheckForUnModeratedPosts("FORUM", Cat_ID, Forum_ID, 0)
end if

' -- Get all the high level(board, category, forum) subscriptions being held by the user
Dim strSubString, strSubArray, strBoardSubs, strCatSubs, strForumSubs, strTopicSubs
if MySubCount > 0 then
	strSubString = PullSubscriptions(0, 0, 0)
	strSubArray  = Split(strSubString,";")
	if uBound(strSubArray) < 0 then
		strBoardSubs = ""
		strCatSubs = ""
		strForumSubs = ""
		strTopicSubs = ""
	else
		strBoardSubs = strSubArray(0)
		strCatSubs = strSubArray(1)
		strForumSubs = strSubArray(2)
		strTopicSubs = strSubArray(3)
	end if
end if

'## Forum_SQL - Find out the Category/Forum status and if it Exists
strSql = "SELECT C.CAT_STATUS, C.CAT_SUBSCRIPTION, " & _ 
	 "C.CAT_MODERATION, C.CAT_NAME, C.CAT_ID, " & _
	 "F.F_STATUS, F.F_SUBSCRIPTION, " & _ 
	 "F.F_MODERATION, F_DEFAULTDAYS, F.F_SUBJECT " & _
	 " FROM " & strTablePrefix & "CATEGORY C, " & _
	 strTablePrefix & "FORUM F " & _
	 " WHERE F.FORUM_ID = " & Forum_ID & _ 
	 " AND C.CAT_ID = F.CAT_ID " & _
	 " AND F.F_TYPE = 0"

set rsCFStatus = my_Conn.Execute (StrSql)

if rsCFStatus.EOF or rsCFStatus.BOF then
	rsCFStatus.close
	set rsCFStatus = nothing
	Response.Redirect("default.asp")
else
	Cat_ID = rsCFStatus("CAT_ID")
	Cat_Name = rsCFStatus("CAT_NAME")
	Cat_Status = rsCFStatus("CAT_STATUS")
	Cat_Subscription = rsCFStatus("CAT_SUBSCRIPTION")
	Cat_Moderation = rsCFStatus("CAT_MODERATION")
	Forum_Status = rsCFStatus("F_STATUS")
	Forum_Subject = rsCFStatus("F_SUBJECT")
	Forum_Subscription = rsCFStatus("F_SUBSCRIPTION")
	Forum_Moderation = rsCFStatus("F_MODERATION")
	if nDays = "" then
		nDays = rsCFStatus("F_DEFAULTDAYS")
	end if
	rsCFStatus.close
	set rsCFStatus = nothing
end if

if strModeration = 1 and Cat_Moderation = 1 and (Forum_Moderation = 1 or Forum_Moderation = 2) then
	Moderation = "Y"
end if
' DEM --> End of Code added for Moderation

if nDays = "" then
	nDays = 30
end if

defDate = DateToStr(dateadd("d",-(nDays),strForumTimeAdjust))

'## Forum_SQL - Get all topics from DB
strSql ="SELECT T.T_STATUS, T.CAT_ID, T.FORUM_ID, T.TOPIC_ID, T.T_VIEW_COUNT, T.T_SUBJECT, " 
strSql = strSql & "T.T_AUTHOR, T.T_STICKY, T.T_REPLIES, T.T_UREPLIES, T.T_LAST_POST, T.T_LAST_POST_AUTHOR, "  
strSql = strSql & "T.T_LAST_POST_REPLY_ID, M.M_NAME, MEMBERS_1.M_NAME AS LAST_POST_AUTHOR_NAME "

strSql2 = " FROM " & strMemberTablePrefix & "MEMBERS M, "
strSql2 = strSql2 & strActivePrefix & "TOPICS T, " 
strSql2 = strSql2 & strMemberTablePrefix & "MEMBERS AS MEMBERS_1 "

strSql3 = " WHERE M.MEMBER_ID = T.T_AUTHOR "
strSql3 = strSql3 & " AND T.T_LAST_POST_AUTHOR = MEMBERS_1.MEMBER_ID "
strSql3 = strSql3 & " AND T.FORUM_ID = " & Forum_ID & " "
if nDays = "-1" then
	if strStickyTopic = "1" then
		strSql3 = strSql3 & " AND (T.T_STATUS <> 0 OR T.T_STICKY = 1)"
	else
		strSql3 = strSql3 & " AND T.T_STATUS <> 0 "
	end if
end if
if nDays > "0" then
	if strStickyTopic = "1" then
		strSql3 = strSql3 & " AND (T.T_LAST_POST > '" & defDate & "' OR T.T_STICKY = 1)"
	else
		strSql3 = strSql3 & " AND T.T_LAST_POST > '" & defDate & "'"
	end if
end if
' DEM --> if not a Moderator, all unapproved posts should not be viewed.
if AdminAllowed = 0 then 
	strSql3 = strSql3 & " AND ((T.T_AUTHOR <> " & MemberID
	strSql3 = strSql3 & " AND T.T_STATUS < "  ' Ignore unapproved/rejected posts
	if Moderation = "Y" then
		strSql3 = strSql3 & "2"  ' Ignore unapproved posts
	else
		strSql3 = strSql3 & "3"  ' Ignore any hold posts
	end if
	strSql3 = strSql3 & ") OR T.T_AUTHOR = " & MemberID & ")"
end if

strSql4 = " ORDER BY"
if strStickyTopic = "1" then
	strSql4 = strSql4 & " T.T_STICKY DESC, "
end if
if strtopicsortfld = "author" then
	strSql4 = strSql4 & " M." & strSortCol & " "
else
	strSql4 = strSql4 & " T." & strSortCol & " "
end if

if strDBType = "mysql" then 'MySql specific code
	if mypage > 1 then 
		intOffset = cLng((mypage-1) * strPageSize)
		strSql5 = strSql5 & " LIMIT " & intOffset & ", " & strPageSize & " "
	end if

	'## Forum_SQL - Get the total pagecount 
	strSql1 = "SELECT COUNT(TOPIC_ID) AS PAGECOUNT "

	set rsCount = my_Conn.Execute(strSql1 & strSql2 & strSql3)
	iPageTotal = rsCount(0).value
	rsCount.close
	set rsCount = nothing

	If iPageTotal > 0 then
		inttotaltopics = iPageTotal
		maxpages = (iPageTotal \ strPageSize )
		if iPageTotal mod strPageSize <> 0 then
			maxpages = maxpages + 1
		end if
		if iPageTotal < (strPageSize + 1) then
			intGetRows = iPageTotal
		elseif (mypage * strPageSize) > iPageTotal then
			intGetRows = strPageSize - ((mypage * strPageSize) - iPageTotal)
		else
			intGetRows = strPageSize
		end if
	else
		iPageTotal = 0
		inttotaltopics = iPageTotal
		maxpages = 0
	end if 

	if iPageTotal > 0 then
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.open strSql & strSql2 & strSql3 & strSql4 & strSql5, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
			arrTopicData = rs.GetRows(intGetRows)
			iTopicCount = UBound(arrTopicData, 2)
		rs.close
		set rs = nothing
	else
		iTopicCount = ""
	end if
 
else 'end MySql specific code

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.cachesize = strPageSize
	rs.open strSql & strSql2 & strSql3 & strSql4, my_Conn, adOpenStatic
		if not rs.EOF then
			rs.movefirst
			rs.pagesize = strPageSize
			inttotaltopics = cLng(rs.recordcount)
			rs.absolutepage = mypage '**
			maxpages = cLng(rs.pagecount)
			arrTopicData = rs.GetRows(strPageSize)
			iTopicCount = UBound(arrTopicData, 2)
		else
			iTopicCount = ""
			inttotaltopics = 0
		end if
	rs.Close
	set rs = nothing
end if

Response.Write	"    <script language=""JavaScript"" type=""text/javascript"">" & vbNewLine & _
		"    <!----- " & vbNewLine & _
		"    function jumpTo(s) {if (s.selectedIndex != 0) location.href = s.options[s.selectedIndex].value;return 1;}" & vbNewLine & _
		vbNewLine & _
		"    function setDays() {document.DaysFilter.submit(); return 0;}" & vbNewLine & _
		"    // -->" & vbNewLine & _
		"    </script>" & vbNewLine

Response.Write	"      <table border=""0"" width=""100%"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td width=""33%"" align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          <a href=""default.asp"">" & getCurrentIcon(strIconFolderOpen,"All Forums","align=""absmiddle""") & "</a>&nbsp;<a href=""default.asp"">All Forums</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","align=""absmiddle""")
if Cat_Status <> 0 then 
	Response.Write	getCurrentIcon(strIconFolderOpen,"","align=""absmiddle""")
else 
	Response.Write	getCurrentIcon(strIconFolderClosed,"","align=""absmiddle""")
end if
Response.Write	"&nbsp;<a href=""default.asp?CAT_ID=" & Cat_ID & """>" & ChkString(Cat_Name,"display") & "</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBlank,"","align=""absmiddle""") & getCurrentIcon(strIconBar,"","align=""absmiddle""")
if ArchiveView = "true" then 
	Response.Write 	getCurrentIcon(strIconFolderArchived,"","align=""absmiddle""")
else
	if Cat_Status <> 0 and Forum_Status <> 0 then 
		Response.Write	getCurrentIcon(strIconFolderOpenTopic,"","align=""absmiddle""")
	else 
		Response.Write	getCurrentIcon(strIconFolderClosedTopic,"","align=""absmiddle""")
	end if
end if
Response.Write	"&nbsp;<a href=""forum.asp?" & ArchiveLink & "FORUM_ID=" & Forum_ID & """>" & ChkString(Forum_Subject,"display") & "</a></font></td>" & vbNewLine & _
		"          <td align=""center"" valign=""bottom"" width=""33%""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine
call PostNewTopic()
Response.Write	"          </font></td>" & vbNewLine & _
		"          <form action=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & ChkString(Request.Querystring,"sqlstring") & """ method=""post"" name=""DaysFilter"">" & vbNewLine & _
		"          <td align=""right"" valign=""bottom"" width=""33%"">" & vbNewLine & _
		"          <select name=""Days"" onchange=""javascript:setDays();"">" & vbNewLine & _
		"          	<option value=""0""" & CheckSelected(ndays,0) & ">Show all topics</option>" & vbNewLine & _
		"          	<option value=""-1""" & CheckSelected(ndays,-1) & ">Show all open topics</option>" & vbNewLine & _
		"          	<option value=""1""" & CheckSelected(ndays,1) & ">Show topics from last day</option>" & vbNewLine & _
		"          	<option value=""2""" & CheckSelected(ndays,2) & ">Show topics from last 2 days</option>" & vbNewLine & _
		"          	<option value=""5""" & CheckSelected(ndays,5) & ">Show topics from last 5 days</option>" & vbNewLine & _
		"          	<option value=""7""" & CheckSelected(ndays,7) & ">Show topics from last 7 days</option>" & vbNewLine & _
		"          	<option value=""14""" & CheckSelected(ndays,14) & ">Show topics from last 14 days</option>" & vbNewLine & _
		"          	<option value=""30""" & CheckSelected(ndays,30) & ">Show topics from last 30 days</option>" & vbNewLine & _
		"          	<option value=""60""" & CheckSelected(ndays,60) & ">Show topics from last 60 days</option>" & vbNewLine & _
		"          	<option value=""120""" & CheckSelected(ndays,120) & ">Show topics from last 120 days</option>" & vbNewLine & _
		"          	<option value=""365""" & CheckSelected(ndays,365) & ">Show topics from the last year</option>" & vbNewLine & _
		"          </select>" & vbNewLine & _
		"          <input type=""hidden"" name=""Cookie"" value=""1"">" & vbNewLine & _
		"          </td></form>" & vbNewLine & _
		"        </tr>" & vbNewLine
if maxpages > 1 then
	Response.Write	"        <tr>" & vbNewLine & _
			"          <td colspan=""3"" align=""right"" valign=""bottom"">" & vbNewLine & _
			"            <table border=""0"" align=""right"">" & vbNewLine & _
			"              <tr>" & vbNewLine
	Call DropDownPaging(1)
	Response.Write	"              </tr>" & vbNewLine & _
			"            </table>" & vbNewLine & _
			"          </td>" & vbNewLine & _
			"        </tr>" & vbNewLine
else
	Response.Write	"        <tr>" & vbNewLine & _
			"          <td colspan=""3""><span style=""font-size: 6px;""><br /></span></td>" & vbNewLine & _
			"        </tr>" & vbNewLine
end if
Response.Write	"      </table>" & vbNewLine & _
		"      <table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
		"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td align=""center"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;</font></b></td>" & vbNewLine & _
		"                <td align=""center"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Topic</font></b></td>" & vbNewLine & _
		"                <td align=""center"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Author</font></b></td>" & vbNewLine & _
		"                <td align=""center"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Replies</font></b></td>" & vbNewLine & _
		"                <td align=""center"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Read</font></b></td>" & vbNewLine & _
		"                <td align=""center"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Last Post</font></b></td>" & vbNewLine
if mlev > 0 or (lcase(strNoCookies) = "1") then 
	Response.Write  "                <td align=""center"" bgcolor=""" & strHeadCellColor & """ nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine
	if (AdminAllowed = 1) then 
		call ForumAdminOptions 
	else
		Response.Write  "                &nbsp;" & vbNewLine
	end if
	Response.Write  "                </font></td>" & vbNewLine
end if 
Response.Write	"              </tr>" & vbNewLine
if iTopicCount = "" then
	Response.Write	"              <tr>" & vbNewLine & _
			"                <td colspan=""7"" bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><b>No Topics Found</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine
else
	tT_STATUS = 0
	tCAT_ID = 1
	tFORUM_ID = 2
	tTOPIC_ID = 3
	tT_VIEW_COUNT = 4
	tT_SUBJECT = 5
	tT_AUTHOR = 6
	tT_STICKY = 7
	tT_REPLIES = 8
	tT_UREPLIES = 9
	tT_LAST_POST = 10
	tT_LAST_POST_AUTHOR = 11
	tT_LAST_POST_REPLY_ID = 12
	tM_NAME = 13
	tLAST_POST_AUTHOR_NAME = 14

	blnStickyTopics = False
	rec = 1
	for iTopic = 0 to iTopicCount
		if (rec = strPageSize + 1) then exit for

		Topic_Status = arrTopicData(tT_STATUS, iTopic)
		Topic_CatID = arrTopicData(tCAT_ID, iTopic)
		Topic_ForumID = arrTopicData(tFORUM_ID, iTopic)
		Topic_ID = arrTopicData(tTOPIC_ID, iTopic)
		Topic_ViewCount = arrTopicData(tT_VIEW_COUNT, iTopic)
		Topic_Subject = arrTopicData(tT_SUBJECT, iTopic)
		Topic_Author = arrTopicData(tT_AUTHOR, iTopic)
		Topic_Sticky = arrTopicData(tT_STICKY, iTopic)
		Topic_Replies = arrTopicData(tT_REPLIES, iTopic)
		Topic_UReplies = arrTopicData(tT_UREPLIES, iTopic)
		Topic_LastPost = arrTopicData(tT_LAST_POST, iTopic)
		Topic_LastPostAuthor = arrTopicData(tT_LAST_POST_AUTHOR, iTopic)
		Topic_LastPostReplyID = arrTopicData(tT_LAST_POST_REPLY_ID, iTopic)
		Topic_MName = arrTopicData(tM_NAME, iTopic)
		Topic_LastPostAuthorName = arrTopicData(tLAST_POST_AUTHOR_NAME, iTopic)

		if AdminAllowed = 1 and Topic_UReplies > 0 then
			Topic_Replies = Topic_Replies + Topic_UReplies
		end if

		If blnStickyTopics = True Then
			If Topic_Sticky and strStickyTopic = "1" Then
				'Do Nothing
			Else
				'We've had sticky topics, and now we are into the regular topics
				Response.Write	"              <tr><td colspan=""7"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Forum Topics</font></b></td></tr>" & vbNewLine
				blnStickyTopics = False
			End If
		End If

		If blnStickyTopics = False Then
			If Topic_Sticky and strStickyTopic = "1" Then
				'We've got sticky topics, so we should draw the line here
				Response.Write	"              <tr><td colspan=""7"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Important Topics</font></b></td></tr>" & vbNewLine
			End If
		End If
		
		Response.Write	"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """ align=""center"" valign=""middle""><a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & """>"
		if Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0 then
			if Topic_Sticky and strStickyTopic = "1" then
				if Topic_LastPost > Session(strCookieURL & "last_here_date") then
					Response.Write	getCurrentIcon(strIconFolderNewSticky,"New Sticky Topic","hspace=""0""")
				else
					Response.Write	getCurrentIcon(strIconFolderSticky,"Sticky Topic","hspace=""0""")
				end if
				blnStickyTopics = True
			else
				' DEM --> Added code for topic moderation
				if Topic_Status = 2 then
					UnApprovedFound = "Y"
					Response.Write 	getCurrentIcon(strIconFolderUnmoderated,"Topic Not Moderated","hspace=""0""") & "</a>" & vbNewline
				elseif Topic_Status = 3 then
					HeldFound = "Y"
					Response.Write 	getCurrentIcon(strIconFolderHold,"Topic on Hold","hspace=""0""") & "</a>" & vbNewline
					' DEM --> end of code Added for topic moderation
				else
					Response.Write	ChkIsNew(Topic_LastPost)
				end if
			end if
		else
			if ArchiveView <> "true" then
				if Cat_Status = 0 then
					strAltText = "Category Locked"
				elseif Forum_Status = 0 then
					strAltText = "Forum Locked"
				else
					strAltText = "Topic Locked"
				end if
			end if
			if ArchiveView = "true" then
				Response.Write	getCurrentIcon(strIconFolderArchived,"Archived Topic","hspace=""0""")
			elseif Topic_LastPost > Session(strCookieURL & "last_here_date") then
				if Topic_Sticky and strStickyTopic = "1" then
					Response.Write	getCurrentIcon(strIconFolderNewStickyLocked,strAltText,"hspace=""0""")
					blnStickyTopics = True
				else
					Response.Write	getCurrentIcon(strIconFolderNewLocked,strAltText,"hspace=""0""")
				end if
			else
				if Topic_Sticky and strStickyTopic = "1" then
					Response.Write	getCurrentIcon(strIconFolderStickyLocked,strAltText,"hspace=""0""")
					blnStickyTopics = True
				else
					Response.Write	getCurrentIcon(strIconFolderLocked,strAltText,"hspace=""0""")
				end if
			end if
		end if
		Response.Write	"</a></td>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """ valign=""middle"" align=""left"">" & vbNewLine & _
				"                <font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>"
		if Topic_Sticky and strStickyTopic = "1" then Response.Write("Sticky:  ")
		Response.Write	"<span class=""spnMessageText""><a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & """>" & ChkString(Topic_Subject,"title") & "</a></span>&nbsp;</font>" & vbNewLine
		if strShowPaging = "1" then 
			Call TopicPaging() 
		end if
		Response.Write	"                </td>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """ valign=""middle"" align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><span class=""spnMessageText"">" & profileLink(chkString(Topic_MName,"display"),Topic_Author) & "</span></font></td>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """ valign=""middle"" align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & Topic_Replies & "</font></td>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """ valign=""middle"" align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & Topic_ViewCount & "</font></td>" & vbNewLine
		if IsNull(Topic_LastPostAuthor) then
			strLastAuthor = ""
		else
			strLastAuthor = "<br />by: <span class=""spnMessageText"">" & profileLink(ChkString(Topic_LastPostAuthorName, "display"),Topic_LastPostAuthor) & "</span>"
			if (strJumpLastPost = "1") then strLastAuthor = strLastAuthor & "&nbsp;" & DoLastPostLink
		end if
		Response.Write	"                <td bgcolor=""" & strForumCellColor & """ valign=""middle"" align=""center"" nowrap><font face=""" & strDefaultFontFace & """ color=""" & strForumFontColor & """ size=""" & strFooterFontSize & """><b>" & ChkDate(Topic_LastPost,"</b>&nbsp;",true) & strLastAuthor & "</font></td>" & vbNewLine
		if mlev > 0 or (lcase(strNoCookies) = "1") then
			Response.Write	"                <td bgcolor=""" & strForumCellColor & """ valign=""middle"" align=""center"" nowrap>" & vbNewLine
			if AdminAllowed = 1 then
				call TopicAdminOptions 
			else
				if Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0 then
					call TopicMemberOptions
				else
					Response.Write	"                &nbsp;" & vbNewLine
				end if
			end if
			Response.Write	"                </td>" & vbNewLine
		end if
		Response.Write	"              </tr>" & vbNewLine

		rec = rec + 1 
	next 
end if
'-------------------------------------------------
' TOPIC SORTING MOD
'-------------------------------------------------
Response.Write	"              <tr>" & vbNewLine & _
		"                <td align=""center"" bgcolor=""" & strHeadCellColor & """ colspan=""6"">" & vbNewLine

dim topicreclow, topicrechigh, topicpage

topicpage = mypage

if (topicpage <= 1) then
	topicreclow = 1
else
	topicreclow = ((topicpage - 1) * strPageSize) + 1
end if

topicrechigh = topicreclow + (rec - 2)

Response.Write	"                <form method=""post"" name=""topicsort"" id=""pagelist"" action=""forum.asp?" & strQStopicsort & """>" & vbNewLine
if ArchiveView = "true" then Response.Write "                <input name=""ARCHIVE"" type=""hidden"" value=""" & ArchiveView & """>" & vbNewLine
Response.Write	"                  <table cellpadding=""0"" cellspacing=""0"" border=""0"" align=""right"" width=""100%"">" & vbNewLine & _
		"                    <tr>" & vbNewLine & _
		"                      <td align=""center"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strHeadFontColor & """>"
if inttotaltopics = 0 then
	Response.Write("No Topics Found")
elseif topicreclow = topicrechigh then
	Response.Write("Showing topic " & topicreclow & " of " & inttotaltopics)
else
	Response.Write("Showing topics " & topicreclow & " - " & topicrechigh & " of " & inttotaltopics)
end if
Response.Write	", sorted by</font></b>&nbsp;<select name=""sortfield"" style=""font-size:10px;"">" & vbNewLine & _
		"                      <option value=""topic""" & CheckSelected(strtopicsortfld,"topic") & ">topic title</option>" & vbNewLine & _
		"                      <option value=""author""" & CheckSelected(strtopicsortfld,"author") & ">topic author</option>" & vbNewLine & _
		"                      <option value=""replies""" & CheckSelected(strtopicsortfld,"replies") & ">number of replies</option>" & vbNewLine & _
		"                      <option value=""views""" & CheckSelected(strtopicsortfld,"views") & ">number of views</option>" & vbNewLine & _
		"                      <option value=""lastpost""" & CheckSelected(strtopicsortfld,"lastpost") & ">last post time</option>" & vbNewLine & _
		"                      </select>&nbsp;<b><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strHeadFontColor & """>in</font></b>&nbsp;<select name=""sortorder"" style=""font-size:10px;"">" & vbNewLine & _
		"                      <option value=""desc""" & CheckSelected(strtopicsortord,"desc") & ">descending</option>" & vbNewLine & _
		"                      <option value=""asc""" & CheckSelected(strtopicsortord,"asc") & ">ascending</option>" & vbNewLine & _
		"                      </select>&nbsp;<b><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strHeadFontcolor & """>order, from</font></b><nobr>&nbsp;<select name=""Days"" style=""font-size:10px;"">" & vbNewLine & _
		"                      <option value=""0""" & CheckSelected(ndays,0) & ">all topics</option>" & vbNewLine & _
		"                      <option value=""-1""" & CheckSelected(ndays,-1) & ">all open topics</option>" & vbNewLine & _
		"                      <option value=""1""" & CheckSelected(ndays,1) & ">the last day</option>" & vbNewLine & _
		"                      <option value=""2""" & CheckSelected(ndays,2) & ">the last 2 days</option>" & vbNewLine & _
		"                      <option value=""5""" & CheckSelected(ndays,5) & ">the last 5 days</option>" & vbNewLine & _
		"                      <option value=""7""" & CheckSelected(ndays,7) & ">the last 7 days</option>" & vbNewLine & _
		"                      <option value=""14""" & CheckSelected(ndays,14) & ">the last 14 days</option>" & vbNewLine & _
		"                      <option value=""30""" & CheckSelected(ndays,30) & ">the last 30 days</option>" & vbNewLine & _
		"                      <option value=""60""" & CheckSelected(ndays,60) & ">the last 60 days</option>" & vbNewLine & _
		"                      <option value=""90""" & CheckSelected(ndays,90) & ">the last 90 days</option>" & vbNewLine & _
		"                      <option value=""120""" & CheckSelected(ndays,120) & ">the last 120 days</option>" & vbNewLine & _
		"                      <option value=""365""" & CheckSelected(ndays,365) & ">the last year</option>" & vbNewLine & _
		"                      </select>" & vbNewLine & _
		"                      <input type=""hidden"" name=""Cookie"" value=""1""><input style=""font-size:10px;"" type=""submit"" name=""Go"" value=""Go""></nobr></td>" & vbNewLine & _
		"                    </tr>" & vbNewLine & _
		"                  </table>" & vbNewLine & _
		"                </form>" & vbNewLine & _
		"                </td>" & vbNewLine
if mLev > 0 or (lcase(strNoCookies) = "1") then
	Response.Write	"                <td align=""center"" bgcolor=""" & strHeadCellColor & """ nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>"
	if (AdminAllowed = 1) then
	    	call ForumAdminOptions
	else
		Response.Write	"                &nbsp;" & vbNewLine
	end if
	Response.Write	"                </font></b></td>" & vbNewLine
end if
Response.Write	"              </tr>" & vbNewLine
'-------------------------------------------------

Response.Write	"            </table>" & vbNewLine & _
		"          </td>" & vbNewLine & _
		"        </tr>" & vbNewLine
if maxpages > 1 then
	Response.Write	"        <tr>" & vbNewLine & _
			"          <td colspan=""7"">" & vbNewLine & _
			"            <table border=""0"" align=""left"">" & vbNewLine & _
			"              <tr>" & vbNewLine
	Call DropDownPaging(2)
	Response.Write	"              </tr>" & vbNewLine & _
			"            </table>" & vbNewLine & _
			"          </td>" & vbNewLine & _
			"        </tr>" & vbNewLine
else
	Response.Write	"        <tr>" & vbNewLine & _
			"          <td colspan=""7""><span style=""font-size: 6px;""><br /></span></td>" & vbNewLine & _
			"        </tr>" & vbNewLine
end if
Response.Write	"      </table>" & vbNewLine & _
		"      <table width=""100%"" align=""center"" border=""0"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td align=""left"" valign=""top"" width=""33%"">" & vbNewLine & _
		"            <table>" & vbNewLine & _
		"              <tr valign=""top"">" & vbNewLine & _
		"                <td valign=""top"" nowrap>" & vbNewLine & _
		"                <p><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>" & vbNewLine & _
		"                " & getCurrentIcon(strIconFolderNew,"New Posts","align=""absmiddle""") & " New posts since last logon.<br />" & vbNewLine & _
		"                " & getCurrentIcon(strIconFolder,"Old Posts","align=""absmiddle""") & " Old Posts."
if lcase(strHotTopic) = "1" then Response.Write	(" (" & getCurrentIcon(strIconFolderHot,"Hot Topic","align=""absmiddle""") & "&nbsp;" & intHotTopicNum & " replies or more.)<br />" & vbNewLine)
Response.Write	"                " & getCurrentIcon(strIconFolderLocked,"Locked Topic","align=""absmiddle""") & " Locked topic.<br />" & vbNewLine
' DEM --> Start of Code added for moderation
if HeldFound = "Y" then
	Response.Write "                " & getCurrentIcon(strIconFolderHold,"Held Topic","align=""absmiddle""") & " Held Topic.<br />" & vbNewline
end if
if UnapprovedFound = "Y" then
	Response.Write "                " & getCurrentIcon(strIconFolderUnmoderated,"UnModerated Topic","align=""absmiddle""") & " UnModerated Topic.<br />" & vbNewline
end if
' DEM --> End of Code added for moderation
Response.Write	"                </font></p></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"            </table>" & vbNewLine & _
		"          </td>" & vbNewLine & _
		"          <td align=""center"" valign=""top"" width=""33%""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine
call PostNewTopic()
Response.Write	"          </font></td>" & vbNewLine & _
		"          <td align=""right"" valign=""top"" width=""33%"" nowrap>" & vbNewLine
%>
<!--#INCLUDE FILE="inc_jump_to.asp" -->
<%
Response.Write	"          </td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine
WriteFooter
Response.End

sub PostNewTopic() 
	if Cat_Status = 0 or Forum_Status = 0 then 
		if (AdminAllowed = 1) then
			Response.Write	"          <a href=""post.asp?method=Topic&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconFolderLocked,"Category Locked","align=""absmiddle""") & "</a>&nbsp;<a href=""post.asp?method=Topic&FORUM_ID=" & Forum_ID & """>New Topic</a><br />" & vbNewLine
		else
			Response.Write	"          " & getCurrentIcon(strIconFolderLocked,"Category Locked","align=""absmiddle""") & "&nbsp;Category Locked<br />" & vbNewLine
		end if
	else 
		if Forum_Status <> 0 then
			Response.Write	"          <a href=""post.asp?method=Topic&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconFolderNewTopic,"New Topic","align=""absmiddle""") & "</a>&nbsp;<a href=""post.asp?method=Topic&FORUM_ID=" & Forum_ID& """>New Topic</a><br />" & vbNewLine
		else
		    	Response.Write	"          " & getCurrentIcon(strIconFolderLocked,"Forum Locked","align=""absmiddle""") & "&nbsp;Forum Locked<br />" & vbNewLine
		end if 
	end if 
	' DEM --> Start of Code added to handle subscription processing.
	if (strSubscription < 4 and strSubscription > 0) and (Cat_Subscription > 0) and Forum_Subscription = 1 and (mlev > 0) and strEmail = 1 then
		Response.Write	"          "
		if InArray(strForumSubs, Forum_ID) then
			Response.Write ShowSubLink ("U", Cat_ID, Forum_ID, 0, "Y") & vbNewLine
		elseif strBoardSubs <> "Y" and not(InArray(strCatSubs,Cat_ID)) then
			Response.Write ShowSubLink ("S", Cat_ID, Forum_ID, 0, "Y") & vbNewLine
		end if
	end if 
	' DEM --> End of code added to handle subscription processing.
end sub

sub ForumAdminOptions() 
	if (AdminAllowed = 1) then 
		if Cat_Status = 0 then 
			if mlev = 4 then
				Response.Write	"                <a href=""JavaScript:openWindow('pop_open.asp?mode=Category&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderUnlocked,"Un-Lock Category","") & "</a>" & vbNewLine
			else
				Response.Write	"                " & getCurrentIcon(strIconFolderLocked,"Category Locked","") & vbNewLine
			end if 
		else 
			if Forum_Status <> 0 then
				Response.Write	"                <a href=""JavaScript:openWindow('pop_lock.asp?mode=Forum&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderLocked,"Lock Forum","") & "</a>" & vbNewLine
			else
				Response.Write	"                <a href=""JavaScript:openWindow('pop_open.asp?mode=Forum&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderUnlocked,"Un-Lock Forum","") & "</a>" & vbNewLine
			end if 
		end if 
		if (Cat_Status <> 0 and Forum_Status <> 0) or (AdminAllowed = 1) then
			Response.Write	"                <a href=""post.asp?method=EditForum&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "&type=0"">" & getCurrentIcon(strIconFolderPencil,"Edit Forum Properties","hspace=""0""") & "</a>" & vbNewLine
		end if 
		if mLev = 4 or (lcase(strNoCookies) = "1") then
			Response.Write	"                <a href=""JavaScript:openWindow('pop_delete.asp?" & ArchiveLink & "mode=Forum&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderDelete,"Delete Forum","") & "</a>" & vbNewLine
			if strArchiveState = "1" then Response.Write("                <a href=""admin_forums.asp?action=archive&id=" & Forum_ID & """>" & getCurrentIcon(strIconFolderArchive,"Archive Forum","") & "</a>" & vbNewLine)
		end if
		' DEM --> Start of Code for Moderated Posting
	        if (UnModeratedPosts > 0) and (AdminAllowed = 1) then
			Response.Write "                <a href=""moderate.asp"">" & getCurrentIcon(strIconFolderModerate,"View All UnModerated Posts","") & "</a>" & vbNewline
	        end if
        	' DEM --> End of Code for Moderated Posting
	end if
end sub

sub DropDownPaging(fnum)
	if maxpages > 1 then
		if mypage = "" then
			pge = 1
		else
			pge = mypage
		end if
		scriptname = request.servervariables("script_name")
		Response.write	"                <form name=""PageNum" & fnum & """ action=""forum.asp"">" & vbNewLine
		Response.Write	"                <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine
		Response.write	"                <input name=""FORUM_ID"" type=""hidden"" value=""" & Forum_ID & """>" & vbNewLine
		Response.write	"                <input name=""sortfield"" type=""hidden"" value=""" & strtopicsortfld & """>" & vbNewLine
		Response.write	"                <input name=""sortorder"" type=""hidden"" value=""" & strtopicsortord & """>" & vbNewLine
		if ArchiveView = "true" then Response.write "                <input name=""ARCHIVE"" type=""hidden"" value=""" & ArchiveView & """>" & vbNewLine
		if fnum = 1 then
			Response.Write("                <b>Page: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & vbNewLine)
		else
			Response.Write("                <b>There are " & maxpages & " Pages of Topics: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & vbNewLine)
		end if
		for counter = 1 to maxpages
			if counter <> cLng(pge) then   
				Response.Write "                	<option value=""" & counter &  """>" & counter & "</option>" & vbNewLine
			else
				Response.Write "                	<option selected value=""" & counter &  """>" & counter & "</option>" & vbNewLine
			end if
		next
		if fnum = 1 then
			Response.Write("                </select><b> of " & maxPages & "</b>" & vbNewLine)
		else
			Response.Write("                </select>" & vbNewLine)
		end if
		Response.Write("                </font></td>" & vbNewLine)
		Response.Write("                </form>" & vbNewLine)
	end if
end sub

sub TopicPaging()
	mxpages = (Topic_Replies / strPageSize)
	if mxPages <> cLng(mxPages) then
		mxpages = int(mxpages) + 1
	end if
	if mxpages > 1 then
		Response.Write("                  <table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine)
		Response.Write("                    <tr>" & vbNewLine)
		Response.Write("                      <td valign=""bottom""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>" & getCurrentIcon(strIconPosticon,"","") & "</font></td>" & vbNewLine)
		for counter = 1 to mxpages
			ref = "                      <td align=""right"" valign=""bottom"" bgcolor=""" & strForumCellColor  & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>" 
			if ((mxpages > 9) and (mxpages > strPageNumberSize)) or ((counter > 9) and (mxpages < strPageNumberSize)) then
				ref = ref & "&nbsp;"
			end if		
			ref = ref & widenum(counter) & "<span class=""spnMessageText""><a href=""topic.asp?"
			ref = ref & ArchiveLink
		        ref = ref & "TOPIC_ID=" & Topic_ID
			ref = ref & "&whichpage=" & counter
			ref = ref & """>" & counter & "</a></span></font></td>"
			Response.Write ref & vbNewLine
			if counter mod strPageNumberSize = 0 and counter < mxpages then
				Response.Write("                    </tr>" & vbNewLine)
				Response.Write("                    <tr>" & vbNewLine)
				Response.Write("                      <td>&nbsp;</td>" & vbNewLine)
			end if
		next				
        Response.Write("                    </tr>" & vbNewLine)
        Response.Write("                  </table>" & vbNewLine)
	end if
end sub

sub TopicAdminOptions() 
	if strStickyTopic = "1" then
		if Topic_Sticky then
			Response.Write	"                <a href=""JavaScript:openWindow('pop_open.asp?mode=STopic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Topic_ForumID & "')"">" & getCurrentIcon(strIconGoDown,"Make Topic Un-Sticky","hspace=""0""") & "</a>" & vbNewLine
		else
			Response.Write	"                <a href=""JavaScript:openWindow('pop_lock.asp?mode=STopic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Topic_ForumID & "')"">" & getCurrentIcon(strIconGoUp,"Make Topic Sticky","hspace=""0""") & "</a>" & vbNewLine
		end if
	end if
	if Cat_Status = 0 then
		Response.Write	"                <a href=""JavaScript:openWindow('pop_open.asp?mode=Category&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Category","hspace=""0""") & "</a>" & vbNewLine
	else
		if Forum_Status = 0 then
			Response.Write	"                <a href=""JavaScript:openWindow('pop_open.asp?mode=Forum&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Forum","hspace=""0""") & "</a>" & vbNewLine
		else 
			if Topic_Status <> 0 then
				Response.Write	"                <a href=""JavaScript:openWindow('pop_lock.asp?mode=Topic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconLock,"Lock Topic","hspace=""0""") & "</a>" & vbNewLine
			else
				Response.Write	"                <a href=""JavaScript:openWindow('pop_open.asp?mode=Topic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Topic","hspace=""0""") & "</a>" & vbNewLine
			end if 
		end if
	end if 
	if (AdminAllowed = 1) or (Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0) then
		Response.Write	"                <a href=""post.asp?" & ArchiveLink & "method=EditTopic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconPencil,"Edit Topic","hspace=""0""") & "</a>" & vbNewLine
	end if
	Response.Write	"                <a href=""JavaScript:openWindow('pop_delete.asp?" & ArchiveLink & "mode=Topic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconTrashcan,"Delete Topic","hspace=""0""") & "</a>" & vbNewLine
	if Topic_Status <= 1 and ArchiveView = "" then
		Response.Write	"                <a href=""post.asp?" & ArchiveLink & "method=Reply&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconReplyTopic,"Reply to Topic","hspace=""0""") & "</a>" & vbNewLine
	end if
	' DEM --> Start of Code for Full Moderation
        if Topic_Status > 1 then
		TopicString = "TOPIC_ID=" & Topic_ID & "&CAT_ID=" & Cat_ID & "&FORUM_ID=" & Forum_ID
               	Response.Write "                <a href=""JavaScript:openWindow('pop_moderate.asp?" & TopicString & "')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject this Topic","hspace=""0""") & "</a>" & vbNewline
        end if
	' DEM --> End of Code for Full Moderation 
 	' DEM --> Start of Code added to handle subscription processing.
	if (strSubscription > 0) and (Cat_Subscription > 0) and Forum_Subscription > 0 and strEmail = 1 then
		Response.Write	"                "
		if InArray(strTopicSubs, Topic_ID) then
			Response.Write ShowSubLink ("U", Cat_ID, Forum_ID, Topic_ID, "N") & vbNewLine
		elseif strBoardSubs <> "Y" and not(InArray(strForumSubs,Forum_ID) or InArray(strCatSubs,Cat_ID)) then
			Response.Write ShowSubLink ("S", Cat_ID, Forum_ID, Topic_ID, "N") & vbNewLine
		end if
	end if 
	' DEM --> End of code added to handle subscription processing.
end sub

sub TopicMemberOptions() 
        if ((Topic_Status > 0 and Topic_Author = MemberID) or (AdminAllowed = 1)) and ArchiveView = "" then
		Response.Write	"                <a href=""post.asp?" & ArchiveLink & "method=EditTopic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconPencil,"Edit Topic","hspace=""0""") & "</a>" & vbNewLine
	end if
        if ((Topic_Status > 0 and Topic_Author = MemberID and Topic_Replies = 0) or (AdminAllowed = 1)) and ArchiveView = "" then
		Response.Write	"                <a href=""JavaScript:openWindow('pop_delete.asp?" & ArchiveLink & "mode=Topic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconTrashcan,"Delete Topic","hspace=""0""") & "</a>" & vbNewLine
	end if
	if Topic_Status <= 1 and ArchiveView = "" then
		Response.Write	"                <a href=""post.asp?" & ArchiveLink & "method=Reply&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconReplyTopic,"Reply to Topic","hspace=""0""") & "</a>" & vbNewLine
 	end if
	' DEM --> Start of Code added to handle subscription processing.
	if (strSubscription > 0) and (Cat_Subscription > 0) and Forum_Subscription > 0 and strEmail = 1 then
		Response.Write	"                "
		if InArray(strTopicSubs, Topic_ID) then
			Response.Write ShowSubLink ("U", Cat_ID, Forum_ID, Topic_ID, "N") & vbNewLine
		elseif strBoardSubs <> "Y" and not(InArray(strForumSubs,Forum_ID) or InArray(strCatSubs,Cat_ID)) then
			Response.Write ShowSubLink ("S", Cat_ID, Forum_ID, Topic_ID, "N") & vbNewLine
		end if
	end if 
	' DEM --> End of code added to handle subscription processing.
end sub

Function DoLastPostLink()
	if Topic_Replies < 1 or Topic_LastPostReplyID = 0 then
		DoLastPostLink = "<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & """>" & getCurrentIcon(strIconLastpost,"Jump to Last Post","align=""absmiddle""") & "</a>"
	elseif Topic_LastPostReplyID <> 0 then
		PageLink = "whichpage=-1&"
		AnchorLink = "&REPLY_ID="
		DoLastPostLink = "<a href=""topic.asp?" & ArchiveLink & PageLink & "TOPIC_ID=" & Topic_ID & AnchorLink & Topic_LastPostReplyID & """>" & getCurrentIcon(strIconLastpost,"Jump to Last Post","align=""absmiddle""") & "</a>"
	else
		DoLastPostLink = ""
	end if
end function


function CheckSelected(chkval1, chkval2)
	if IsNumeric(chkval1) then chkval1 = cLng(chkval1)
	if (chkval1 = chkval2) then
		CheckSelected = " selected"
	else
		CheckSelected = ""
	end if
end function
%>
