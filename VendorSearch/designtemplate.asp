<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">

<%
'-------------------- MBELCHER: DYNAMICALLY LOAD JQUERY/JAVASCRIPT/CSS : ------------------------------
'This makes the files load no matter what folder we are in instea of jsut the root. Example: \\10.18.12.41\Sites\ifa\custom\fraud\directory_fraud.asp
strHTTP_REFERER = Request.ServerVariables("HTTP_REFERER")
strSERVER_NAME = Request.ServerVariables("SERVER_NAME")
strConsolefoldername = trim(lcase(application("consolefoldername")))
strSiteDomainName = trim(Application("siteurl"))
strPAGENAME = Request.ServerVariables("URL")
strBootstrapMinCSS_URL = "http://" & strSERVER_NAME & "/" & "css/bootstrap.min.css"
strBootstrapMinCSS_CDN = "http://" & strSERVER_NAME & "/" & "/netdna.bootstrapcdn.com/bootstrap/3.3.2/css/bootstrap.min.css"
strCstyleCSS = "http://" & strSERVER_NAME & "/" & "css/cstyle.css"
strMember_pagesCSS = "http://" & strSERVER_NAME & "/" & "css/member_pages.css"
strInstallCSS = "http://" & strSERVER_NAME & "/" & "css/install.css"
strsmstyleCSS = "http://" & strSERVER_NAME & "/" & "css/smstyle.css?v=0"
strfont_awesomeMinCSSCDN = "http://" & strSERVER_NAME & "/" & "/maxcdn.bootstrapcdn.com/font-awesome/4.4.0/css/font-awesome.min.css"
strfont_awesomeMinCSS = "http://" & strSERVER_NAME & "/" & "css/font-awesome.min.css"
strVendorSearchCSS = "http://" & strSERVER_NAME & "/" & "css/vendor_styles.css"

strmodernizrCustomJS = "http://" & strSERVER_NAME & "/" & "js/modernizr.custom.js"
strBootstrapMinJS = "http://" & strSERVER_NAME & "/" & "js/bootstrap.min.js"
strResponsiveSlidesJS = "http://" & strSERVER_NAME & "/" & "js/responsiveslides.js"
strjQueryRotateJS = "http://" & strSERVER_NAME & "/" & "jQueryRotate.js"
strSidenav_mobileJS = "http://" & strSERVER_NAME & "/" & "js/sidenav_mobile.js"
strjBootstrapHoverDropdownMinJS = "http://" & strSERVER_NAME & "/" & "js/bootstrap-hover-dropdown.min.js"

'Response.Write ("<b>SERVER_NAME:</b> " & strSERVER_NAME & "<br>")
'Response.Write ("<b>strConsolefoldername:</b> " & strConsolefoldername & "<br>")
'Response.Write ("<b>siteurl:</b> " & trim(Application("siteurl")) & "<br>")
'Response.Write ("<b>strPAGENAME:</b> " & strPAGENAME & "<br>")
'Response.Write ("<b>strBootstrapMinCSS_URL:</b> " & strBootstrapMinCSS_URL & "<br>")
'-------------------------------------------------------------------------------------------------------------------------          
%>







<%

Call MetaTags()

%>
<title>
<%
Call ThePageTitle()
%></title>

    <!-- Le styles -->
    <!-- MBELCHER: CASE: 25487 we have files located in different folders so need to dynamically link to these in order to use same design template -->   

	<link href="<%=strBootstrapMinCSS_URL %>" rel="stylesheet" type="text/css">
	<link href="<%=strBootstrapMinCSS_CDN %>" rel="stylesheet">

    <link href="<%=strCstyleCSS %>" rel="stylesheet">
    <link href="<%=strMember_pagesCSS %>" rel="stylesheet" type="text/css">
    <link href="<%=strInstallCSS %>" rel="stylesheet" type="text/css">
    
    <link href="<%=strsmstyleCSS %>" rel="stylesheet" type="text/css">
    <link rel="stylesheet" href="<%=strfont_awesomeMinCSSCDN %>">
    <link href="<%=strfont_awesomeMinCSS  %>" rel="stylesheet" type="text/css">
    <link href="<%=strVendorSearchCSS %>" rel="stylesheet" type="text/css">

    <script src="<%=strmodernizrCustomJS %>"></script>
    
    <!--<script src="js/jquery-1.10.2.min.js"></script> duplicate-->
	<!--<script src="js/lightbox-2.6.min.js"></script> in tlbase now-->
    <!--<link href="css/lightbox.css" rel="stylesheet" />in tlbase now-->
    
    	<!-- HTML5 shim, for IE6-8 support of HTML5 elements -->
		<!--[if lt IE 9]>
		<script src="js/html5shiv.js"></script>
		<![endif]-->
		<!--[if  IE 8]>
		<style>
		.navbar .main-menu ul li:hover ul {
		display: block;
		}
		</style>
		<![endif]-->

<%
Call HeadCode()

Call Javascripts()

%>

	</head>
    
	<body>

		<div class="container">
			
            <header>
            	<div class="header-top">
                	<div class="inner-wrapper">	
                	<section>
                		<div class="header-top-right">
                		<div class="row">
                			<div class="col-xs-12 col-sm-12 col-md-12">
                				<div class="search-img">
                					<a href="/search.asp?type=basic"><img src="../images/theme/search-img.png" alt="logo"/></a>
                				</div>

                				<div class="member-login">
                					<%sa_login%>
                				</div>	
                			</div><!--col-12-ends-->	
                		</div><!--row-ends-->
                		</div><!--header-top-right-ends-->
                	</section>
                	</div><!--inner-wrapper-ends-->
                </div><!--header-top-ends-->	

                		<div class="clear"></div>

                <div class="inner-wrapper">		
                	<section>
                		<div class="header-bottom">
                		<div class="row">
                			<div class="col-xs-12 col-sm-4 col-md-4">
	            				<div class="logo-img">
									<a href="/index.asp"><img src="../images/theme/logo.png" alt="logo"/></a>
								</div>
            				</div><!--col-6-ends-->

                			<div class="col-xs-12 col-sm-8 col-md-8">
								<div class="nav-bg">		
                        			<%call show_nav (1,2,"sm_asca_primary", "") %>
                        		</div><!--nav-bg-ends-->	
                        	</div>
                    	</div><!--row-ends-->
                    	</div><!--header-bottom-ends-->
                	</section>
                </div><!--inner-wrapper-ends-->
            </header>     
                    
                    <div class="clear"></div>
					
			<div class="inner-wrapper">  

				<section>		
					<div class="main-content-subpage">
						<div class="row">
							<div class="col-xs-12 col-sm-3 col-md-3">
								<div class="subpage-content-left">

								<div class="mobile-container">
									<div class="mobile-menu"><a href="#menu"><h6>Sub Menu</h6><i class="fa fa-caret-down fa-lg"></i></a>
                                	</div>       	
									<div class="subpage-menu">
                                             
                                            <!--
												
                                                <ul>
													<li>
														<a href="#">Subnavigation 1</a>
													</li>
													<li>
														<a href="#">Subnavigation 2</a>
														<ul>
															<li>
																<a href="#">Third Level One</a>
															</li>
															<li class="active">
																<a href="#">Third Level One</a>
															</li>
															<li>
																<a class="remove-border" href="#">Third Level One</a>
															</li>
														</ul>
													</li>
													<li>
														<a href="#">Subnavigation 3</a>
													</li>
													<li>
														<a href="#">Subnavigation 4</a>
													</li>
												</ul>
											</div>
                                            -->
                                            
                                            
                                            
                                            	<%
                                            if nav_oncontent=1 then
                                           
										    call show_nav (2,2,"sm_asca_secondary", "") 
                                           
										    end if
                                            if instr(lcase(request.ServerVariables("script_name")), "store_")>0 then
                                                store_sidebar
                                            end if
											%>
                                            
                                            <%		
                                            If Len(session("sa_id")) > 0 and showmembermenu()=true Then 'showmembermenu function located in crumbs.asp
					                            AF_Menu 
					                        End If
                                            %>

                                        </div><!--subpage-menu-ends-->
                                     </div><!--mobile-container-->    
                                            
                                           	<div class="title-section">
												<%showcmn "Subpage Left Content" %>
											</div>

										</div><!--subpage-content-left-ends-->
									</div><!--col-3-ends-->

									<div class="col-xs-12 col-sm-9 col-md-9">
										<div class="subpage-content-right">
											<div class="breadcrumb-menu">
												<div class="breadcrumb-inner">
													<%call displaylistcrumbs("")%>
												</div>
											</div>
											<div class="subpage-content">
                                            
                                            <!--------------------START PAGE BODY---------------------->
										  	  <%CALL PageBody()%>
                                            <!---------------------END PAGE BODY---------------------->
												
											</div>
										</div>
									</div>
						</div>				
					</div>
				</section>
			</div><!--inner-wrapper-ends-->  
				
					  <div class="clear"></div>
                    
				<footer>
					<div class="inner-wrapper">
						<div class="row">
							<div class="col-xs-12 col-sm-8 col-md-8">
								<div class="footer-address">
			                    	<% showcmn "Global Footer" %>
			                    </div><!--global-footer-ends-->	
								<!-- <span class="copy">&copy; <%=datepart("yyyy",now())%> New York State Association of Counties</span> -->					
							</div><!--col-8-ends-->	
							
							<div class="col-xs-12 col-sm-4 col-md-4">
								<div class="footer-logo">
									<a href="http://www.membershipsoftware.org" target="_blank">Powered by: Timberlake</a>
								</div>
							</div><!--col-4-ends-->	
						</div><!--row-ends-->
                   	</div><!--inner-wrapper-ends-->
				</footer>

		</div><!--/.fluid-container-->

<div id="divMBELCHER">
<%'MBELCHER: will be commented out in production. 

%>
</div>

		<!-- Le javascript
		================================================== -->

        <!--#include file="../tlbase/common/jquery/jquery.asp"--> 
      	<script src="<%=strBootstrapMinJS %>"></script>
		<script src="<%=strResponsiveSlidesJS %>"></script>
		<script src="<%=strjQueryRotateJS %>"></script>
		<script src="<%=strSidenav_mobileJS %>"></script>
		<script src="<%=strjBootstrapHoverDropdownMinJS %>"></script>

		<script type="text/javascript">
		 	$(document).ready(function () {
		 		$('.dropdown-toggle').dropdownHover();
				$('#Basic_category').change(function(){
					var optionSelected = $("option:selected",this).val();
					if(optionSelected === "8"){
						$('#toggleText').show();
						$('select[name=Basic_state]').prop('disabled',true);
					}
					else
					{
						$('#toggleText').hide();
						$('select[name=Basic_state]').prop('disabled',false);
					}
				});
		 });
		 
		</script>

	</body>
</html>
<!--#include file="includes/google_analytics.asp"-->




















