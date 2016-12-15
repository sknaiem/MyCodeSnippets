

</div> <!-- /oldtables -->
</div> <!-- /container -->


<div id="divMBELCHER">
<%'MBELCHER: will be commented out in production. 
  'Response.Write ("<br><b>PAGENAME: </b> " & strPAGENAME & "<br>")
  'Response.Write ("<br><b>IFA_fraudAdd_JS: </b> " & strIFA_fraudAdd_JS & "<br>")

  'strScriptLoaded = "NOTHING AT ALL"
%>
</div>

		<!-- Le javascript
		================================================== -->

 

<!--#include file="../../tlbase/common/jquery/jquery.asp"-->
<!--#include file="../../tlbase/common/jquery/jqueryUI.asp"-->
<!--#include file="../../tlbase/common/bootstrap/bootstrap.asp"--> 


        <%'MBELCHER: Dynamic loading of scripts for one page only %> 
        <% If strPAGENAME = "/ifa/custom/admin/admin_Fraud.asp" Then%> 
        <!--#include file="../../tlbase/common/jquery/jqueryDataTables.asp"--> 
        <% End If %>    

		<script type="text/javascript">
		 	$(document).ready(function () {

		 	    //MBELCHER: CASE: 25487 Watchlist Fraud Search
		 	    <% If (strPAGENAME = "/ifa/custom/admin/admin_Fraud.asp") AND (process = "") Then %>
		        $.getScript("<%=strIFA_fraudSearch %>");
		 	    <% strScriptLoaded = "strIFA_fraudSearch" %>
		        <% End If %>

		 	    //MBELCHER: CASE: 25487 Watchlist Fraud Delete
		 	    <% If (strPAGENAME = "/ifa/custom/admin/admin_Fraud_Delete.asp") AND (process = "") Then %>
		        $.getScript("<%=strIFA_fraudDelete %>");
		 	    <% strScriptLoaded = "strIFA_fraudDelete" %>
		        <% End If %>

		 	    //MBELCHER: CASE: 25487 Watchlist Fraud Edit
		 	    <% If (strPAGENAME = "/ifa/custom/admin/admin_Fraud_Edit.asp") AND (process = "") Then %>
		        $.getScript("<%=strIFA_fraudEdit %>");
		 	    <% strScriptLoaded = "strIFA_fraudEdit" %>
		        <% End If %>


		 	    //if (window.jQuery) {
		 	    // jQuery is loaded  
		 	    //alert("jQuery is loaded <br>");
		 	    //} else {
		 	    // jQuery is not loaded
		 	    //alert("jQuery is not loaded<br>");
		 	    //}


		 });
		</script>


<% 'Response.Write ("<b>strScriptLoaded: </b>" & strScriptLoaded & " is loaded<br><br><br>") %>


<footer>
    <div class="container">
        <div class="col-lg-8">
            <span class="copyright">&copy; <%=Year(Now())%> Timberlake Membership Software<br />
            <a href="<%=strBaseServerURL %>" target="_blank">Public Site</a>&nbsp;     <%if session("username")="TimberlakeAdmin" then%> |&nbsp; <a href="<%=strCustomNavBaseLink%>/splash_admin.asp" target="main">Admin</a>&nbsp;<%end if %> |&nbsp; <a href="<%=strCustomNavBaseLink%>/logout.asp" target="_top">Sign Out</a>&nbsp; |&nbsp; 

            <a href="<%=strBaseServerURL %>/tlbase/tops/zendesk/zendesk.asp" target="_blank">Support</a>

            <!-- Fail safe -->
            <!-- <a href="http://support.membershipsoftware.org" target="_blank">Request Support</a> -->

            <!-- <a href="../tlbase/secure/api/zendeskSSO/back_up_saml_sso_zd.asp" target="_blank">Request Support</a> -->
            
            </span>
        </div>
        <div class="col-lg-4">
            <a href="http://membershipsoftware.org" target="_blank"><img class="footerlogo" src="<%=strBaseServerURL %>/images/theme/footerlogo.png" /></a>
        </div>
    </div>
</footer>
