

<%  
   
'-------------------- MBELCHER: DYNAMICALLY LOAD JQUERY/JAVASCRIPT/CSS/ And Custom control : ------------------------------
'This makes the files load no matter what folder we are in instead of only the root. Example: \\10.18.12.41\Sites\ifa\custom\fraud\directory_fraud.asp
mbFolderPath = "../../console"
strHTTP_REFERER = Request.ServerVariables("HTTP_REFERER")
strSERVER_NAME = Request.ServerVariables("SERVER_NAME")
strConsolefoldername = trim(lcase(application("consolefoldername")))
strSiteDomainName = trim(Application("siteurl"))

'---------- Menu and Link Control ----------
strAF_MenuCustom = ""' ON = we are in the client custom folder 

'strCustomFolder = "/custom/fraud"' 
'strCustomURL = "http://" & strSERVER_NAME & strCustomFolder & "/"
strPAGENAME = Request.ServerVariables("URL")
strCustomURL = "http://" & strSERVER_NAME & "/" & strConsolefoldername & "/console" ' & "/"
strBaseServerURL = "http://" & strSERVER_NAME

'Solves user menu Link issue when we are in a custom folder.
'if we are in a custom folder we then set custom on.
'This could be integrated into the console version of TL_header at some point if needed to fix up links for folders outside console.
strSERVER_URL = Request.ServerVariables("URL")
If InStr(strSERVER_URL, "/custom") > 0 Then
    strAF_MenuCustom = "ON"
    strCustomNavBaseLink = strCustomURL
    Else
    strAF_MenuCustom = "OFF" 
    strCustomNavBaseLink = ""
End If

'CUSTOM
strjqPagination = "http://" & strSERVER_NAME & "/" & strConsolefoldername & "/tlbase/common/jquery/jqPagination/js/jquery.jqpagination.min.js"
strIFA_fraudAdd_JS = "http://" & strSERVER_NAME & "/" & strConsolefoldername & "/tlbase/common/jquery/custom/IFA/IFA_fraudAdd.js"
strIFA_fraudSearch = "http://" & strSERVER_NAME & "/" & strConsolefoldername & "/tlbase/common/jquery/custom/IFA/IFA_fraudSearchAdmin.js"
strIFA_fraudDelete = "http://" & strSERVER_NAME & "/" & strConsolefoldername & "/tlbase/common/jquery/custom/IFA/IFA_fraudDelete.js"
strIFA_fraudEdit = "http://" & strSERVER_NAME & "/" & strConsolefoldername & "/tlbase/common/jquery/custom/IFA/IFA_fraudEdit.js"
'-------------------------------------------

'Response.Write ("<b>URL:</b> " & strSERVER_URL & "<br>")
'Response.Write ("<b>strAF_MenuCustom:</b> " & strAF_MenuCustom & "<br>")
'Response.Write ("<b>strCustomNavBaseLink :</b> " & strCustomNavBaseLink  & "<br>")
'Response.Write ("<b>strPAGENAME:</b> " & strPAGENAME & "<br>")
'-------------------------------------------------------------------------------------------------------------------------          
 %>
 
<!-- Bootstrap core CSS -->
<link href="<%=strBaseServerURL %>/css/bootstrap.css" rel="stylesheet">
<link href="<%=strBaseServerURL %>/css/bootstrap-theme.min.css" rel="stylesheet">
<link href="<%=strBaseServerURL %>/css/timberlake.css" rel="stylesheet">

<link href='http://fonts.googleapis.com/css?family=Open+Sans:400,300,700' rel='stylesheet' type='text/css'>

<div class="container">

    <section>
    
    <div class="page-header">
        <div class="col-md-3">
            <div class="logo"><a href="http://www.membershipsoftware.org" target="_blank"><img src="<%=strBaseServerURL %>/images/theme/logo.png" alt="Timberlake" /></a></div>
        </div>
        <div class="col-sm-6 col-md-3">
            <div class="hello">
                <h2><!-- @DEV, data-notifications are for future development - <a data-notifications="3" class="notify"> --><a href="<%=strCustomNavBaseLink %>/modifyaccount.asp">Hello <%if trim(session("firstname")&" ")<>"" then%><%=trim(session("firstname"))%><%else%><%=trim(session("username"))%><%end if %>!</a></h2>
                <div class="hellodate">Today is <%=FormatDateTime(now(),1)%></div>
            </div>
        </div>
        <div class="col-sm-6 col-md-6">
            <div class="toplinks">
                <ul>
                    <li><a href="<%=strBaseServerURL %>/tlbase/tops/zendesk/zendesk.asp" target="_blank">Support</a></li>
                    <!-- Fail safe -->
                    <!-- <li><a href="http://support.membershipsoftware.org" target="_blank">Request Support</a></li> -->
                    <!-- @DEV, class="notice" - "notice" is for future development -->  <!-- <a href="http://training.membershipsoftware.org/Release/Release_Notes.pdf">Release Notes</a> --> <!-- to link to PDF, just change href to PDF's location -->
                    <li><a href="<%=strCustomNavBaseLink %>/newstarted.asp?start=Y&navclick=213">Setup</a></li>
                    <li><a  href="<%=trim(application("siteurl"))%>" target="_newbrowser">Public Site</a></li>
                    <li><a href="<%=strCustomNavBaseLink %>/logout.asp" class="signout">Sign Out</a></li>
                </ul>
            </div>
            <div class="search">
                <form name="search" action="" method="post" onSubmit="javascript:go();">
                    <span>Quick Search</span>
                    <select name="searchtype" size="1" id="Select1" align="absmiddle">                                          
                         <option value="web" >Website Content</option>                  
                         <option value="member" >Membership Content</option>
                    </select> 
                    <input align="middle" type="text" value="" border="1px" size="50" name="keyword" >
                    <button type="submit" value="Go"><img src="/images/theme/search.png" alt=""></button>       
                 </form>
                 <script language="JavaScript" type="text/JavaScript">
                
                    function go()
                    {
                    //alert(document.search.searchtype);
                    if (document.search.searchtype.value=="web")
                    {
                        document.search.action = "<%=strCustomNavBaseLink %>/sb_manage.asp?n=y&parent=0&keyword=" + document.search.keyword.value;
                    document.search.submit();
                    }
                    else if (document.search.searchtype.value=="member")
                    {
                        document.search.action = "<%=strCustomNavBaseLink %>/AF_ContactsList.asp?keyword=" + document.search.keyword.value;
                    document.search.submit();
                    }
                    }
                    </script>
            </div>
        </div>
    
    <div class="clear"></div>
    
    <div class="row">
    <!-- static navbar -->
          <div class="navbar navbar-default" role="navigation">
                        <!-- Brand and toggle get grouped for better mobile display -->
                        <div class="navbar-header">
                            <button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-ex1-collapse">
                                <span class="sr-only">Toggle navigation</span>
                                <span class="icon-bar"></span>
                                <span class="icon-bar"></span>
                                <span class="icon-bar"></span>
                            </button>
                        </div>

                        <div class="collapse navbar-collapse navbar-ex1-collapse">
                            <ul class="nav navbar-nav">
                             <%dim navdrs
       counter=1
            set navdrs=getrecordset("SELECT * From Console_Categories where active=1 order by priority")
            do until navdrs.eof
            if navdrs("permissions")<>"" then
                usepermissions=trim(navdrs("permissions"))
            else
                usepermissions=""
            end if%>
                            
                                <li class="menu-item dropdown"><%if navdrs("disabled") and trim(navdrs("disabledurl"))<>"" and not isnull(navdrs("disabledurl")) Then 

                                %><a href="<%=replace(trim(navdrs("disabledurl")), "../../../console/", "")%>" class="dropdown-toggle"><%else %><a href="<%=addstr%><%=replace(trim(navdrs("introlink")), "../../../console/", "")%>" class="dropdown-toggle"><%end if %><%=trim(navdrs("shortname"))%></a>
                            <%if navdrs("disabled") then
                                'do nothing

                            else%>

                                    <ul class="dropdown-menu">
                                    <%set ors=getrecordset("SELECT * From console_options where categoryid="&navdrs("id")&" and active=1 order by priority") %>
                                    <%do until ors.eof%>
                                     <li class="menu-item "><a href="<%=strCustomNavBaseLink%>/<%=replace(trim(ors("link")), "console/", "")%><%if instr(ors("link"), "?") then%>&navclick=<%=ors("id")%><%else %>?navclick=<%=ors("id")%><%end if%>"><%=trim(ors("Name"))%></a></li>                
      <%                                  ors.movenext
loop%>
                                    </ul>
                                    <%end if %>
                                </li>
                               <%  counter=counter+1
            navdrs.movenext
            loop%>
          



                            </ul>
                        </div>
                </div> <!-- end navbar -->
            </div>
                
            <div class="row">
                <ul class="topscrumbs">
                    <%if trim(metars("crumb1")&" ")<>"" then %>
                        <li><a href="<%=strCustomNavBaseLink%>/<%=trim(metars("crumb1url")&" ") %>"><%=trim(metars("crumb1")&" ")%></a></li>
                    <%
                    lastpg=trim(metars("crumb1")&" ")
                    elseif trim(session("crumbtitle1")&" ")="" then
                        'do nothing
                    else %>
                         <li><a href="<%=strCustomNavBaseLink%>/<%=replace(trim(session("crumblink1")&" "), "console/", "") %>"><%=trim(session("crumbtitle1")&" ")%></a></li>
                    <%
                    lastpg=trim(session("crumbtitle1")&" ")
                    end if %>
                    <%if trim(metars("crumb2")&" ")<>"" then %>
                        <li><a href="<%=strCustomNavBaseLink%>/<%=replace(trim(metars("crumb2url")&" "), "console/", "") %>"><%=trim(metars("crumb2")&" ")%></a></li>
                    <%
                    lastpg=trim(metars("crumb2")&" ")
                     elseif trim(session("crumbtitle2")&" ")="" then
                        'do nothing
                    else %>
                         <li><a href="<%=strCustomNavBaseLink%>/<%=replace(trim(session("crumblink2")&" "), "console/", "") %>"><%=trim(session("crumbtitle2")&" ")%></a></li>
                    <%
                    lastpg=trim(session("crumbtitle2")&" ")
                    end if %>
                    <%if trim(metars("crumb3")&" ")<>"" then %>
                        <li><a href="<%=strCustomNavBaseLink%>/<%=replace(trim(metars("crumb3url")&" "), "console/", "") %>"><%=trim(metars("crumb3")&" ")%></a></li>
                    <%
                    lastpg=trim(metars("crumb3")&" ")
                    end if %>
                   
                    <%if TOPS_PAGE_TITLE<>lastpg then %>
                    <li><a href="#"><%=TOPS_PAGE_TITLE%></a></li>
                    <%end if %>
                </ul>
            </div>
            
        </div><!-- /page-header -->
                
    </section>

    <%if session("username")="TimberlakeAdmin" and session("metaedit") then 
     
    %>
    <%
   
    %>
    <iframe src="<%=strCustomNavBaseLink %>/metadataform.asp?id=<%=TOPS_PAGE_ID%>" frameborder="0" width="1100" height="225"></iframe>
    
    

    <%end if %>



    <div class="oldtables">




















