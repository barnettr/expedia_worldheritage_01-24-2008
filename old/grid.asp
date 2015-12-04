<%
' Grab varible values from QueryString
Dim paramRegion: paramRegion = Replace(Request("region"),"_"," ")
Dim paramState: paramState = Request("state")
Dim paramUrlPage: paramUrlPage = Request.ServerVariables("URL") 
Dim paramCurrentPage: paramCurrentPage = Request("page")
Dim paramPageSize: paramPageSize = Request("pview")
Dim paramSortCol: paramSortCol = Request("sortCol")
Dim paramSortDir: paramSortDir = Request("sortDir")

' Set default values if they do not exist in QueryString
If (paramRegion = "") Then paramRegion = "North America" End If
If (paramState = "") Then paramState = "" End If
If (paramCurrentPage = "") Then paramCurrentPage = "1" End If
If (paramPageSize = "") Then paramPageSize = "25" End If
If (paramSortCol = "") Then paramSortCol = "state" End If
If (paramSortDir = "") Then paramSortDir = "ascending" End If
Dim paramSortType: paramSortType = "text"
Dim paramMaxPages: paramMaxPages = "5"
Dim paramURL: paramURL = Request.ServerVariables("URL")
Dim paramRegionFile: paramRegionFile = "worldheritage_"&Replace(paramRegion," ","_")&".xml"


'------------------------------------
' Function to read XML file
'------------------------------------
'used to load xml file against the dropdown xslt file
Function loadDDL(strXMLFile, strXSLFile)
    Dim myRender
    ' creating an object of XMLDOM
    Set xsldoc = Server.CreateObject("MSXML2.FreeThreadedDOMDocument") 
    xsldoc.async = False
    xsldoc.load strXSLFile

    Set myTemplate = Server.CreateObject("MSXML2.XSLTemplate")
    myTemplate.stylesheet = xsldoc

    Set xmldoc = Server.CreateObject("MSXML2.DOMDocument") 
    xmldoc.async = False
    xmldoc.load strXMLFile

    Set myProc = myTemplate.createProcessor
    myProc.input = xmldoc
	If paramRegion <> "" Then
		myProc.addParameter "region",paramRegion
		myProc.addParameter "state",paramState
		myProc.addParameter "url",paramURL
		myProc.addParameter "pview",paramPageSize		
	End If
    myProc.transform
    myRender = myProc.output
    'response
    Response.Write myRender

    Set xmldoc = Nothing
    Set xsldoc = Nothing
End Function

Function loadGrid(strXMLFile, strXSLFile)
	Dim myRender

	Set xsldoc = Server.CreateObject("MSXML2.FreeThreadedDOMDocument") 
	xsldoc.async = False
	xsldoc.load strXSLFile

	Set myTemplate = Server.CreateObject("MSXML2.XSLTemplate")
	myTemplate.stylesheet = xsldoc

	Set xmldoc = Server.CreateObject("MSXML2.DOMDocument") 
	xmldoc.async = False
	xmldoc.load strXMLFile
	Set myProc = myTemplate.createProcessor
	myProc.input = xmldoc
	If paramRegion <> "" Then
		myProc.addParameter "region",paramRegion
		myProc.addParameter "state",paramState
		myProc.addParameter "page",paramCurrentPage
		myProc.addParameter "pview",paramPageSize
		myProc.addParameter "maxPages",paramMaxPages
		myProc.addParameter "SortCol",paramSortCol
		myProc.addParameter "SortDir",paramSortDir 
		myProc.addParameter "SortType",paramSortType
	End If
	myProc.transform
	myRender = myProc.output

	Response.Write myRender

	Set xmldoc = Nothing
	Set xsldoc = Nothing
End Function
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>World Heritage Testing Page</title>
<style>
span.counter, div.counter {
     padding: 1px 3px 1px 3px;
     margin: 2px;
     border: 1px solid #EEE;
     font-weight: bold;
     background-color: #4985c7;
     color: #FFF;
     }
</style>
<link title="combined" rel="stylesheet" type="text/css" media="screen" href="mapview.css" />
</head>
<body>

<script src="http://dev.virtualearth.net/mapcontrol/mapcontrol.ashx?v=6"></script>
<script type="text/javascript" src="getBrowser.js"></script>
<script type="text/javascript" src="getCustomPushpin.js"></script>
<script type="text/javascript" src="xmlProcessor.js"></script>
<script type="text/javascript" src="setMapStyle.js"></script>

<form id="Criteria" method="get" action="">
    <input type="hidden" id="sRegion" name="sRegion" value="<%= paramRegion %>">
	<input type="hidden" id="sFile" name="sFile" value="<%= paramRegionFile %>">
    <input type="hidden" id="sState" name="sState" value="<%= paramState %>">
    <input type="hidden" id="paramCurrentPage" name="paramCurrentPage" value="<%= paramCurrentPage %>">
    <input type="hidden" id="paramPageSize" name="paramPageSize" value="<%= paramPageSize %>">
    <input type="hidden" id="paramSortCol" name="paramSortCol" value="<%= paramSortCol %>">
    <input type="hidden" id="paramSortDir" name="paramSortDir" value="<%= paramSortDir %>">
	<input type="hidden" id="paramSortType" name="paramSortType" value="<%= paramSortType %>">
</form>

<table border="0" cellpadding="0" cellspacing="0" id="wrapper" style="background-image:url(HP_Page_bkgd.jpg);">
   <tr>
        <td height="50"></td>
   </tr>
    <tr>
        <td width="350" style="padding-left:10px; padding-right:10px; background-color:#faf9fb;">
            <table border="0" cellpadding="2" cellspacing="0" width="100%">
                <tr>
                    <td width="175">
                        <select onChange="window.open(this.options[this.selectedIndex].value,'_top');">
	                        <option value="<%=paramURL%>?region=Africa" <% If paramRegion = "Africa" Then Response.Write "selected=""selected""" End If %>>Africa</option>
	                        <option value="<%=paramURL%>?region=Arab_States" <% If paramRegion = "Arab States" Then Response.Write "selected=""selected""" End If %>>Arab States</option>
	                        <option value="<%=paramURL%>?region=Asia" <% If paramRegion = "Asia" Then Response.Write "selected=""selected""" End If %>>Asia</option>
	                        <option value="<%=paramURL%>?region=Caribbean" <% If paramRegion = "Caribbean" Then Response.Write "selected=""selected""" End If %>>Caribbean</option>
	                        <option value="<%=paramURL%>?region=Europe" <% If paramRegion = "Europe" Then Response.Write "selected=""selected""" End If %>>Europe</option>
	                        <option value="<%=paramURL%>?region=Latin_America" <% If paramRegion = "Latin America" Then Response.Write "selected=""selected""" End If %>>Latin America</option>
	                        <option value="<%=paramURL%>?region=North_America" <% If paramRegion = "North America" Then Response.Write "selected=""selected""" End If %>>North America</option>
	                        <option value="<%=paramURL%>?region=Pacific" <% If paramRegion = "Pacific" Then Response.Write "selected=""selected""" End If %>>Pacific</option>
                        </select>
                    </td>
                    <td width="175">
                        <% loadDDL server.MapPath("worldheritage_"&Replace(paramRegion," ","_")&".xml"), server.MapPath("worldheritage_ddl.xslt") %>                    
                    </td>
                </tr>
            </table>
        </td>
        <td width="100%">&nbsp;
        </td>
    </tr> 
    <tr>
        <td width="350" valign="top" style="padding-left:10px; padding-right:10px; padding-bottom:10px; background-color:#faf9fb;">
            <% loadGrid server.MapPath(paramRegionFile), server.MapPath("worldheritage_grid.xslt") %>
        </td>
        <td width="100%" valign="top" style="padding-top:0px; padding-left:20px;">
            <table border="0" cellpadding="0" cellspacing="0" id="map">
                <tr>
                    <td width="100%" style="padding-left:6px; padding-bottom:10px;">
                        <div id="showhide" style="display:block; padding:1px 1px 1px 1px; margin:0px; border:1px solid #e6e7e7; font-weight:bold; background-color:#f4f9f9; color:#000000; text-decoration:none; font-family:Arial; font-size:8px; position:absolute; left:537px; top:92px; z-index:30;"><a href="javascript:void(0)" style="text-decoration:none;" onclick="HideDashboard('hide');"><img src="hide.gif" id="hide" border="0" width="10" height="10" style="width:10px; height:10px; display:inline;" /></a></div>
                        <div id="hideshow" style="display:none; padding:1px 1px 1px 1px; margin:0px; border:1px solid #e6e7e7; font-weight:bold; background-color:#f4f9f9; color:#00156e; text-decoration:none; font-family:verdana; font-size:11px; position:absolute; left:0px; top:42px; z-index:30;"><a href="javascript:void(0)" onclick="ShowDashboard('show');"><img src="show.gif" id="show" border="0" width="10" height="10" style="display:inline;" /></a></div>
                        <div id="expedia_logo" style="position:absolute; top:404px; width:93px; height:29px; left:626px; z-index:30; filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(src='logo.png', sizingMethod='scale');"><img src="logo.png" id="logo" border="0" width="93" height="29" style="display:none; position:absolute; bottom:2px; left:8px" /></div>
                        <div id="inprogress_img" style="display:none;" align="center"><p style="vertical-align:middle; padding-top:2px; padding-bottom:2px; font-weight:bold;" id="p_holder"><img src="mozilla_blu.gif" id="loader" align="absmiddle" style="width:16px; height:16px; border:0px;" /> <strong>Loading . . .</strong></p></div>
                        <div id="mapContainer" style="width:500px; margin-top:0px; margin-right:0px; margin-left: 0px; background-color:; border-left:1px solid #b4c6d8; border-right:1px solid #b4c6d8; border-bottom:1px solid #b4c6d8;">
                            <div id="myMap" style="position:relative; width:650px; height:350px; border-top:1px solid #b4c6d8;">
                                
                            </div>
                            <div style="margin-bottom:0px; width:650px; height:29px; background:transparent url(http://media.expedia.com/media/content/expus/graphics/other/wizard_span.gif) repeat-x top left;overflow:hidden;">
                                <form name="styleForm" style="margin:0px 0px 0px 0px;">
                                <table border="0" cellspacing="0" cellpadding="0" style="width:100%;">
                                    <tr>
                                        <td align="center" style="width:33%; padding-top:3px; font-family:arial,sans-serif; color:#333; font-size: 11px;" valign="middle">
                                            <input type="radio" name="styleType" onclick="SetMapStyle();"  /> Aerial
                                        </td>
                                        <td align="center" style="width:33%; padding-top:3px; font-family:arial,sans-serif; color:#333; font-size: 11px;" valign="middle">
                                            <input type="radio" onclick="SetMapStyle();" name="styleType" checked="checked" /> Road
                                        </td>
                                        <td align="center" style="width:33%; padding-top:3px; font-family:arial,sans-serif; color:#333; font-size: 11px;" valign="middle">
                                            <input type="radio" onclick="SetMapStyle();" name="styleType" /> Hybrid
                                        </td>
                                    </tr>
                                </table>
                                </form>
                            </div>
                        </div>
                    </td>
                </tr>
            </table>                    
        </td>
    </tr>
    <tr>
        <td height="200"></td>
    </tr>
</table>
<br /><br />

</body>
<script type="text/javascript">
    getBrowser();
//function stopError() {
//  return true;
//}
//window.onerror = stopError;



</script>
</html>
