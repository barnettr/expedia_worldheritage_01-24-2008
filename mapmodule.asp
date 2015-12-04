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
If (paramPageSize = "") Then paramPageSize = "10" End If
If (paramSortCol = "") Then paramSortCol = "state" End If
If (paramSortDir = "") Then paramSortDir = "ascending" End If
Dim paramSortType: paramSortType = "text"
Dim paramMaxPages: paramMaxPages = "999"
Dim paramURL: paramURL = Request.ServerVariables("URL")


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
<html>
<head>

<% 
    If InStr(1, Request.ServerVariables("HTTP_USER_AGENT"), "Firefox/") Then

%>
	<link title="combined" rel="stylesheet" type="text/css" media="screen" href="mapview_ff.css" />
<%
    Else
%>
    <link title="combined" rel="stylesheet" type="text/css" media="screen" href="mapview.css" />
<% 
    End If
%>

<title>World Heritage</title>
<meta http-equiv="Content-Language" content="en-us" />
<meta http-equiv="Content-type" content="text/html; charset=iso-8859-1"/>
<script type="text/javascript" src="http://dev.virtualearth.net/mapcontrol/mapcontrol.ashx?v=6"></script>
<script type="text/javascript" src="zoom.js"></script>
<script type="text/javascript">

var level = "";
var pinID = 1;
var selStyle = "r";  // a, h, r
var selMode = VEMapMode.Mode2D;
var fixed = 0;
var showSwitch = 1
var zoom = "";
var map = null;
var arrPin = new Array();
var pinid = 0;
var panningpin = null;
var panningBehavior = false;
var panningDelay = null;
var nextZ = 1000;
var showInfoboxFromGrid = false;
var state = "<%= paramState %>";
var region = "<%= paramRegion %>"

if (state == "") {
  <% If Request("region") = "" Then %>
  zoom = 1;
  <% Else %>
  zoom = 2;
  <% End If %>
}else {
  displayZoomLevel(stateArray,state);  
}

function GetMap()
{
	map = new VEMap('myMap');
	var defaultMapLoc = GetDefaultMapLoc();
	map.SetDashboardSize(VEDashboardSize.Normal);	
	map.onLoadMap = MapLoaded;
	map.LoadMap(defaultMapLoc,zoom,'r');
	//map.ClearInfoBoxStyles();
	map.AttachEvent("onendpan", OnMapMoved);
	map.AttachEvent("onchangeview", OnMapMoved);
	map.AttachEvent("onmouseover", mouseoverPin);
	map.AttachEvent("onmouseout", mouseoutPin);
	map.AttachEvent("onendzoom", showCurrentZoom);	
    map.AttachEvent("onmousewheel", onscrollwheel);	
	AddPushpins();
}

function onscrollwheel(e)
{
     return true;
} 

function MapLoaded()
{
    document.getElementById('preloader').style.display='none';
	document.getElementById('myGrid').style.display='block';
	<% If Request("region") <> "" Then %>	
	document.getElementById('totallistitems').innerText=numTotalListItems;
	document.getElementById('totallistitemslabel').innerText=numTotalListItemsLabel;
	<% End If %>
	showCurrentZoom();
}

function GetDefaultMapLoc()
{
	var latLon = new VELatLong(arrPin[0].lat,arrPin[0].lng);
	return latLon;
}

function AddPushpins()
{
	for (x = 0; x < arrPin.length; x++)
	{
		var pin = BuildPushpin(arrPin[x].id,arrPin[x].region,arrPin[x].site,arrPin[x].location,arrPin[x].state,arrPin[x].suitcase,arrPin[x].flag,arrPin[x].desc,arrPin[x].img,arrPin[x].lat,arrPin[x].lng,arrPin[x].locid)
		map.AddShape(pin);
		arrPin[x].pinid = pin.GetID();
	}
}

function BuildPushpin(id,region,site,location,state,suitcase,flag,desc,img,lat,lng,locid)
{
	var loc = new VELatLong(lat,lng);
	var title = "<table border='0' cellpadding='0' cellspacing='0' style='width:100%'><tr><td align='left' valign='top' style='width:16px;'><span style='padding: 1px 3px 1px 3px; margin: 2px; border: 1px solid #EEE; font-weight: bold; background-color: #4985c7; color: #FFF;'>" + id + "</span></td><td align='left' valign='top' style='text-align:left; padding-left:6px;'><a href='' style='text-decoration:none; border:0px; font-weight:bold; text-align:left; color:#4985c7'>" + site + "</a></td><td align='right' valign='top' style='width:41px; height:34px; padding-left:4px;'><img src='" + suitcase + "' style='border:none;'></td></tr></table>"
	
	var description = "<div style='margin-left:0;'><table border='0' cellpadding='0' cellspacing='0'><tr><td align='right' valign='top' style='padding-right:0px; padding-top:10px;'><img src='" + img + "' style='border: solid 1px #c0c0c1; padding:1px 1px 1px 1px; background-color:#ffffff;' align='left' width='88'></td><td align='left' valign='top' style='font-family:arial; font-size:11; color:#333333; padding-left:6px; padding-top:7px; text-align:left;'><a href='javascript:this.map.HideInfoBox();'><img src='http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_close.gif' align='right' style='margin-left:3px;' id='close'></a><a href='sites.asp?region="+region+"&locid="+locid+"' class='infoBoxLinkOff' onmouseover='this.className=\"infoBoxLinkOn\";' onmouseout='this.className=\"infoBoxLinkOff\";' target='_top'><b>" + site + "</b></a><div style='margin-top:3px;margin-bottom:3px;'>"+ location +"<br /><div style='margin-top:2px;'>"+ state +"</div></div>" + truncate(desc,region,locid) + "<br /><br /><a href='sites.asp?region="+region+"&locid="+locid+"' target='_top'><img src='http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_landing_arrow.gif' style='border:none;padding-right:2px;'></a> <a href='sites.asp?region="+region+"&locid="+locid+"' class='sitepageLinkOff' onmouseover='this.className=\"sitepageLinkOn\";' onmouseout='this.className=\"sitepageLinkOff\";' style='vertical-align:top;' target='_top'>Explore this site</a></td></tr></table></div>";
	if (navigator.userAgent.indexOf("Firefox") != -1)
	{
	    var description = "<div style='margin-left:0;'><table border='0' cellpadding='0' cellspacing='0'><tr><td align='right' valign='top' style='padding-top:10px;'><img src='" + img + "' style='border: solid 1px #c0c0c1; padding:1px 1px 1px 1px; background-color:#ffffff;' align='left' width='88'></td><td align='left' valign='top' style='font-family:arial; font-size:11px; color:#333333; padding-left:6px; padding-top:7px; padding-right:8px; text-align:left;'><a href='javascript:this.map.HideInfoBox();'><img src='http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_close.gif' align='right' style='margin-left:3px;padding-right:6px;'></a><a href='sites.asp?region="+region+"&locid="+locid+"' class='infoBoxLinkOff' onmouseover='this.className=\"infoBoxLinkOn\";' onmouseout='this.className=\"infoBoxLinkOff\";' target='_top'><b>" + site + "</b></a><div style='margin-top:3px;margin-bottom:3px;'>"+ location +"<br /><div style='margin-top:2px;'>"+ state +"</div></div>" + truncate(desc,region,locid) + "<br /><br /><a href='sites.asp?region="+region+"&locid="+locid+"' target='_top'><img src='http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_landing_arrow.gif' style='border:none;padding-right:2px;'></a> <a href='sites.asp?region="+region+"&locid="+locid+"' class='sitepageLinkOff' onmouseover='this.className=\"sitepageLinkOn\";' onmouseout='this.className=\"sitepageLinkOff\";' style='vertical-align:top;' target='_top'>Explore this site</a></td></tr></table></div>";
	}
	if (navigator.userAgent.indexOf("MSIE 7") != -1)
    {
        var description = "<div style='margin-left:0;'><table border='0' cellpadding='0' cellspacing='0'><tr><td align='right' valign='top' style='padding-right:0px; padding-top:10px;'><img src='" + img + "' style='border: solid 1px #c0c0c1; padding:1px 1px 1px 1px; background-color:#ffffff;' align='left' width='88'></td><td align='left' valign='top' style='font-family:arial; font-size:11px; color:#333333; padding-left:6px; padding-top:7px; text-align:left;'><a href='javascript:this.map.HideInfoBox();'><img src='http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_close.gif' align='right' style='margin-left:3px;padding-right:6px;'></a><a href='sites.asp?region="+region+"&locid="+locid+"' class='infoBoxLinkOff' onmouseover='this.className=\"infoBoxLinkOn\";' onmouseout='this.className=\"infoBoxLinkOff\";' target='_top'><b>" + site + "</b></a><div style='margin-top:3px;margin-bottom:3px;'>"+ location +"<br /><div style='margin-top:2px;'>"+ state +"</div></div>" + truncate(desc,region,locid) + "<br /><br /><a href='sites.asp?region="+region+"&locid="+locid+"' target='_top'><img src='http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_landing_arrow.gif' style='border:none;padding-right:2px;'></a> <a href='sites.asp?region="+region+"&locid="+locid+"' class='sitepageLinkOff' onmouseover='this.className=\"sitepageLinkOn\";' onmouseout='this.className=\"sitepageLinkOff\";' style='vertical-align:top;' target='_top'>Explore this site</a></td></tr></table></div>";
    }
	var customIcon = "<div class='pinImgOff' onmouseover='this.className=\"pinImgOn\";' onmouseout='this.className=\"pinImgOff\";'>" + id + "</div>";
	
	var shape = new VEShape(VEShapeType.Pushpin, loc);
	shape.SetCustomIcon(customIcon);
	shape.SetTitle(title);
	shape.SetDescription(description);
	map.ClearInfoBoxStyles();
	return shape;
}

function mouseoverPin(e)
{
	if (e.elementID != null)
    {
		var row = getRowID(map.GetShapeByID(e.elementID).GetID());
		showRow(row,'pin');
	}
}

function mouseoutPin(e)
{
	if (e.elementID != null)
    {
		var row = getRowID(map.GetShapeByID(e.elementID).GetID());
		hideRow(row);
	}
}

function GridShowInfoBox(obj)
{
	var pin = map.GetShapeByID(obj);
	pin.SetZIndex(2000);
	if (pin)
	{
		var shapeLat = Math.round(pin.GetIconAnchor().Latitude);
		var shapeLng = Math.round(pin.GetIconAnchor().Longitude);
		var mapLat = Math.round(map.GetCenter().Latitude);
		var mapLng = Math.round(map.GetCenter().Longitude);
		if(shapeLat == mapLat && shapeLng == mapLng)
		{
		    if (showInfoboxFromGrid == true)
		    {
			    map.ShowInfoBox(pin);
			}
		}
		else
		{
			panningpin = pin;
			panningDelay = setTimeout("CenterPin()",1000);
		}
	}
}

function GridHideInfoBox(obj)
{
    clearTimeout(panningDelay);
	var pin = map.GetShapeByID(obj);
	pin.SetZIndex(1000);
	map.HideInfoBox(pin);
}

function CenterPin()
{
	panningBehavior = true;
	map.SetCenter(panningpin.GetIconAnchor());
}

function OnMapMoved()
{
	if (panningBehavior)
	{
	    if (showInfoboxFromGrid == true)
	    {
		    setTimeout("map.ShowInfoBox(panningpin)",10);
		}
	}
	panningBehavior = false;
}

function getPinID(val)
{
	for (x = 0; x < arrPin.length; x++)
	{
		if (arrPin[x].id == val)
		{
			return arrPin[x].pinid;
		}
	}
}

function getRowID(val)
{
	for (x = 0; x < arrPin.length; x++)
	{
		if (arrPin[x].pinid == val)
		{
			return arrPin[x].id;
		}
	}
}

function showRow(id,type)
{
	var row = 'row_'+id;
	document.getElementById(row).className = 'gridRowHover';
	
	var gridIcon = 'gridIcon_'+id;
	document.getElementById(gridIcon).className = 'gridIconOff';
	
	var pin = map.GetShapeByID(getPinID(id));
	if(type=='pin')
	{
		//document.getElementById(row).scrollIntoView(true);
		SetZIndex(pin,1);
	}
	else
	{
	    var customIcon = "<div class='pinImgOn' onmouseover='this.className=\"pinImgOn\";' onmouseout='this.className=\"pinImgOff\";'>" + id + "</div>";
	    pin.SetCustomIcon(customIcon);
		pin.SetZIndex(2000);
	}
}

function hideRow(id,type)
{
	var row = 'row_'+id;
	document.getElementById(row).className = 'gridRow';
	
	var gridIcon = 'gridIcon_'+id;
	document.getElementById(gridIcon).className = 'gridIcon';	

	var pin = map.GetShapeByID(getPinID(id));
	if(type=='pin')
	{
		SetZIndex(pin,-1);
	}
	else
	{
	    var customIcon = "<div class='pinImgOff' onmouseover='this.className=\"pinImgOn\";' onmouseout='this.className=\"pinImgOff\";'>" + id + "</div>";
	    pin.SetCustomIcon(customIcon);	
		pin.SetZIndex(1000);
	}
}

function swapIcon(pin,id,state)
{
    var icon = "<div style='width:28px;height:32px;background-image:url(http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_pinicon_"+state+".gif);background-repeat:no-repeat;color: #000;text-align:center;font-size:11px;font-weight:bold; padding-top: 4px;'>"+id+"</div>";
	pin.SetCustomIcon(icon);   
}

function SetZIndex(shape,delta)
{
	if (shape && shape.GetPrimitive)
	{
		var shapeElem = document.getElementById(shape.GetPrimitive(0).iid);
		if (shapeElem && shapeElem.style)
		{
			shapeElem.style.zIndex = parseInt(shapeElem.style.zIndex) + delta;
		}
	}
}
function truncate(txt,region,locid)
{
	var len = 160;
	var p = txt;
	if (p)
	{
		var trunc = p;
		if (trunc.length > len)
		{
			trunc = trunc.substring(0, len);
			trunc = trunc.replace(/\w+$/, '');
			trunc += '... <a href="sites.asp?region='+region+'&locid='+locid+'" class="sitepageLinkOff" onmouseover="this.className=\'sitepageLinkOn\';" onmouseout="this.className=\'sitepageLinkOff\';" style="text-decoration:underline;" target="_top">More info</a> <img src="http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_arrow-sm.gif">';
			return trunc;
		}
	}
}

function showCurrentZoom()
{
	//document.getElementById("zoom").innerText=map.GetZoomLevel(); 
	//document.title="Zoom Level: "+map.GetZoomLevel();
}

function setFindType(obj)
{
    if(obj.value == "featured")
    {
        document.location="mapmodule.asp";
    }
    else
    {
        document.location="mapmodule.asp?region=<%=paramRegion%>";
    }
}

function addEvent(obj,evType,fn){if(obj.addEventListener){obj.addEventListener(evType,fn,false);return true;}else if(obj.attachEvent){var r=obj.attachEvent("on"+evType,fn);return r;}else{return false;}}

addEvent(window, 'load', GetMap);

</script>
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0" width="736" id="filterBox" style="background-image:url(http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_quickfacts_lg_v3.gif); background-repeat:no-repeat;">
    <tr>
        <td valign="top" width="736" style="padding:10px 20px 20px 20px;">
            <table border="0" cellpadding="0" cellspacing="0">
                <tr>
                    <td width="100%" class="explaination">
                        <div style="padding-bottom:8px;">Use our interactive map to zero-in on World Heritage sites around the world. Simple browse by destination or site name to move the map, pinpointing the location of a site. Search results are also displayed below.</div>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td valign="top" style="padding:0px 20px 15px 20px; margin-top:0px;" width="736">
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr>
                    <td valign="top" class="buttons" width="235">
                        <table border="0" cellpadding="0" cellspacing="0">
                            <tr>
                                <td valign="top">
                                    <span><input onclick="setFindType(this);" type="radio" id="radFindType" name="radFindType" value="featured" class="" <% If Request("region") = "" Then Response.Write "checked=""checked""" End If %> /></span>
                                </td>
                                <td valign="top" style="padding-left:4px; padding-top:2px; font-size:11px;">
                                    <span style="white-space:nowrap;"><strong>Browse featured World Heritage sites</strong></span>
                                    <div style="margin-top:2px;">Reasearch your sustainable tourism options from the Serengeti to the Great Barrier Reef</div>       
                                </td>
                            </tr>
                        </table>
                    </td>
                    <td align="center" class="spacer"><img src="http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_spacer.gif" width="1" height="70" /></td>
                    <td valign="top" class="buttons">
                        <table border="0" cellpadding="0" cellspacing="0">
                            <tr>
                                <td valign="top" width="200">
                                    <span><input onclick="setFindType(this);" type="radio" id="radFindType" name="radFindType" value="region" class="" <% If Request("region") <> "" Then Response.Write "checked=""checked""" End If %> /></span>
                                </td>
                                <td valign="top" style="padding-left:4px;font-size:11px; padding-top:2px;">
                                    <span style="white-space:nowrap;"><strong>Find World Heritage sites in:</strong></span>
                                    <div style="margin-bottom:5px; margin-top:5px;">
                                        <select id="ddlRegion" onChange="document.location=this.options[this.selectedIndex].value;" style="width:150px; font-size:11px;">
                                            <option value="<%=paramURL%>?region=Africa" <% If paramRegion = "Africa" Then Response.Write "selected=""selected""" End If %>>Africa</option>
	                                        <option value="<%=paramURL%>?region=Arab_States" <% If paramRegion = "Arab States" Then Response.Write "selected=""selected""" End If %>>Arab States</option>
	                                        <option value="<%=paramURL%>?region=Asia" <% If paramRegion = "Asia" Then Response.Write "selected=""selected""" End If %>>Asia</option>
	                                        <option value="<%=paramURL%>?region=Caribbean" <% If paramRegion = "Caribbean" Then Response.Write "selected=""selected""" End If %>>Caribbean</option>
	                                        <option value="<%=paramURL%>?region=Europe" <% If paramRegion = "Europe" Then Response.Write "selected=""selected""" End If %>>Europe</option>
	                                        <option value="<%=paramURL%>?region=Latin_America" <% If paramRegion = "Latin America" Then Response.Write "selected=""selected""" End If %>>Latin America</option>
	                                        <option value="<%=paramURL%>?region=North_America" <% If paramRegion = "North America" Then Response.Write "selected=""selected""" End If %>>North America</option>
	                                        <option value="<%=paramURL%>?region=Pacific" <% If paramRegion = "Pacific" Then Response.Write "selected=""selected""" End If %>>Pacific</option>
                                        </select>
                                    </div>
                                    <div>
                                        <% loadDDL server.MapPath("worldheritage_"&Replace(paramRegion," ","_")&".xml"), server.MapPath("worldheritage_ddl.xslt") %>
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </td>
                    <td align="center" class="spacer"><img src="http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_spacer.gif" width="1" height="70" /></td>
                    <td valign="top" class="buttons">
                        <table border="0" cellpadding="0" cellspacing="0">
                            <tr>
                                <td valign="top" style="padding-left:4px;font-size:11px;">
                                    <div style="margin-bottom:5px;"><strong>Search for sites by name:</strong></div>
                                    <!-- BEGIN AUTOCOMPLETE -->
                                    <script type="text/javascript" src="autocomplete.js"></script>
                                    <style type="text/css">
                                    div.autosuggest
                                    {
                                        font-family: arial, helvetica, sans serif;
                                        font-size: 11px;
                                        font-weight: normal;
	                                    position: absolute;
	                                    padding: 0;
	                                    margin-top: 5px;
	                                    margin-left: 0px;
	                                    height: 150px;
	                                    overflow-y: scroll;
	                                    background: #ffffff;
	                                    width: 172px;
	                                    border: 1px solid #CCCCCC;
                                    }

                                    html>body div.autosuggest
                                    {
                                        width: 188px;
	                                    margin-top: 5px;
	                                    margin-left: 0px;
	                                    z-index: 1000;
                                    }

                                    div.autosuggest ul
                                    {
	                                    list-style: none;
	                                    margin: 0;
	                                    padding: 0;
	                                    overflow: hidden;
	                                    background-color: #ffffff; /* offstate background color of items */
	                                    position: relative;
	                                    top: 0px;
                                    }

                                    html>body div.autosuggest ul
                                    {
                                        padding: 0;
                                    }

                                    div.autosuggest ul li
                                    {
	                                    color: #ccc;
	                                    padding: 0;
	                                    margin: 0;
	                                    text-align: left;
	                                    line-height: 10px;	
                                    }

                                    div.autosuggest ul li a
                                    {
	                                    color: #4985c7; /* site text not including input text */
	                                    display: block;
	                                    text-decoration: none;
	                                    background-color: transparent;
	                                    position: relative;
	                                    padding: 0;
	                                    width: 100%;
                                    }

                                    div.autosuggest ul li a:hover
                                    {
	                                    background-color: #FFF9DD; /* hover state if highlighted item is different */
                                    }

                                    div.autosuggest ul li.as_highlight a:hover
                                    {
	                                    background-color: #FFF9DD; /* hover over site front */
                                    }

                                    div.autosuggest ul li.as_highlight span
                                    {
	                                    background-color: #FFF9DD; /* hover over site front */
	                                    text-decoration: underline;
                                    }

                                    div.autosuggest ul li a span
                                    {
	                                    display: block;
	                                    padding: 2px;
	                                    font-weight: normal;
	                                    line-height: 14px;
                                    }

                                    div.autosuggest ul li a small
                                    {
                                        padding: 2px;
	                                    font-weight: normal;
	                                    color: #666666; /* info text */
                                    }

                                    div.autosuggest ul li.as_highlight a small
                                    {
	                                    color: #666666; /* hover info text */
                                    }

                                    div.autosuggest ul li.as_highlight a
                                    {
	                                    color: #4985c7; /* hover site text not including input text */
	                                    background-color: #FFF9DD; /* hover over site behind */
                                    }

                                    div.autosuggest ul li.as_warning
                                    {
                                        width: 172px;
	                                    font-weight: normal;
	                                    text-align: center;
	                                    color: #333333; /* no sites text */
                                    }

                                    div.autosuggest ul em
                                    {
	                                    font-style: normal;
	                                    color: #CC6600; /* input text */
                                    }
                                    </style>
                                    <!--[if IE 7]>
                                    <style type="text/css">
                                    html>body div.autosuggest
                                    {
                                        width: 190px;
	                                    margin-top: 5px;
	                                    margin-left: 0px;
                                    }
                                    </style>
                                    <![endif]-->
                                    <input type="text" id="whsearch" style="width: 185px;" value="" />
                                    <script type="text/javascript">
                                    var autocomplete = {
	                                    script: function (input) { return "autocompleteService.asp?json=false&input="+input },
	                                    json: false,
	                                    varname:"input",
	                                    noresults:"No sites located",
	                                    minchars:3,
	                                    cache:false,
	                                    callback: function (obj) { top.window.location="sites.asp?region="+obj.regn+"&locid="+obj.id; }
                                    };
                                    var as_xml = new bsn.AutoSuggest('whsearch', autocomplete);
                                    </script>                                    
                                    <!-- END AUTOCOMPLETE -->
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<br />
<div>
<table width="728" cellpadding="2" cellspacing="0" border="0">
    <tr>
        <td bgcolor="#ffffff" class="wrapper">
            <table width="726" cellpadding="0" cellspacing="0" border="0">
                <tr>
                    <td valign="top" width="413" bgcolor="#ffffff" class="mapWrapper">
                        <table width="413" cellpadding="1" cellspacing="0" border="0">
                            <tr>
                                <td width="413">
                                    <table width="413" cellpadding="0" cellspacing="0" border="0">
                                        <tr>
                                            <td width="413" bgcolor="#ffffff" style="">
                                                <table width="411" cellpadding="0" cellspacing="2" border="0">
                                                    <tr>
                                                        <td align="left" valign="top">
                                                            <table cellpadding="0" cellspacing="0" border="0">
                                                                <tr>
                                                                    <td valign="top">
                                                                        <table width="411" cellpadding="0" cellspacing="0" border="0">
                                                                            <tr>
                                                                                <td width="411" valign="top">
                                                                                    <table width="411" cellpadding="0" cellspacing="0" border="0">
                                                                                        <tr>
                                                                                            <td valign="top" width="411" class="mapHeadline">
                                                                                                    <% If Request("region") <> "" Then %>
                                                                                                    <span id="totallistitems"></span> <span id="totallistitemslabel"></span>&nbsp;
                                                                                                    <% End If %>
                                                                                                    <div id="mapContainer" style="width:409px; height:435px; border:solid 1px #cccccc; margin-top:0px;">
                                                                                                        <div id="preloader"><img src="http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_preloader.gif" align="absmiddle" alt="Loading map..." />&nbsp;Loading map...</div>
                                                                                                        <div id="myMap" style="position:relative;width:409px;height:433px;margin:1px 1px 1px 1px;">
                                                                                                            
                                                                                                        </div>
                                                                                                        
                                                                                                    </div>
                                                                                            </td>
                                                                                        </tr>
                                                                                    </table>
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                        
                    </td>
                    <td width="300" valign="top" bgcolor="#ffffff" class="gridWrapper">
                        <table width="300" cellpadding="0" cellspacing="0" border="0">
                            <tr>
                                <td width="295" valign="top">
                                    <table width="295" cellpadding="0" cellspacing="0" border="0">
                                        <tr>
                                            <td width="295" valign="top">
                                                <table width="100%" cellpadding="0" cellspacing="0" border="0">
                                                    <tr>
                                                        <td valign="top" width="100%">
                                                            
                                                           <div id="myGrid">
                                                                <%
                                                                If (Request("region") = "") Then
                                                                   loadGrid server.MapPath("worldheritage_featured.xml"), server.MapPath("worldheritage_featuredgrid.xslt") 
                                                                Else
                                                                    loadGrid server.MapPath("worldheritage_"&Replace(paramRegion," ","_")&".xml"), server.MapPath("worldheritage_grid.xslt")
                                                                End If
                                                                %>
                                                            </div> 
                                                            
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
</div>
<script language="javascript" type="text/javascript">
if (navigator.userAgent.indexOf("MSIE 7") != -1)
{
    document.getElementById("myMap").style.cssText = "position:relative;width:407px;height:433px;margin:1px 8px 8px 1px;";
    document.getElementById("filterBox").style.cssText = "background-image:url(http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_quickfacts_lg_v3.gif); background-repeat:no-repeat;";
}
if (navigator.userAgent.indexOf("Firefox") != -1)
{
    document.getElementById("myMap").style.cssText = "position:relative;width:407px;height:433px;margin:1px 8px 8px 1px;";
    document.getElementById("filterBox").style.cssText = "background-image:url(http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_quickfacts_lg_v3.gif); background-repeat:no-repeat; height:154px;";
}
</script>
</body>
</html>
