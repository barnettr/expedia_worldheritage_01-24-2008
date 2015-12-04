<%
Dim region: region = Replace(Request("region")," ","_")
Dim locid: locid = Request("locid")

If (region = "" OR locid = "") Then
    Response.Redirect("find.asp")
End If    

Set xmldoc = Server.CreateObject("MSXML2.DOMDocument") 
xmldoc.async = false
xmldoc.Load(Server.MapPath("worldheritage_"&region&".xml"))

strSelectNode = "//worldheritage/row[@id=" & locid & "]/"

Set objId = xmldoc.selectNodes(strSelectNode & "@id")
If objId.length > 0 Then numId = objId(0).Text Else numId = "" End If
Set objLat = xmldoc.selectNodes(strSelectNode & "@lat")
If objLat.length > 0 Then numLat = objLat(0).Text Else numLat = "" End If
Set objLong = xmldoc.selectNodes(strSelectNode & "@long")
If objLong.length > 0 Then numLong = objLong(0).Text Else numLong = "" End If
Set objInscribeDate = xmldoc.selectNodes(strSelectNode & "inscribeDate")
If objInscribeDate.length > 0 Then dtInscribeDate = objInscribeDate(0).Text Else dtInscribeDate = "" End If
Set objFlagCode = xmldoc.selectNodes(strSelectNode & "code")
If objFlagCode.length > 0 Then strFlagCode = objFlagCode(0).Text Else strFlagCode = "" End If
Set objRegion = xmldoc.selectNodes(strSelectNode & "region")
If objRegion.length > 0 Then strRegion = objRegion(0).Text Else strRegion = "" End If
Set objState = xmldoc.selectNodes(strSelectNode & "state")
If objState.length > 0 Then strState = objState(0).Text Else strState = "" End If
Set objLocations = xmldoc.selectNodes(strSelectNode & "location")
If objLocations.length > 0 Then
    for i = 0 to (objLocations.length-1)
        strLocations = strLocations & objLocations(i).Text
        If (objLocations(i).Text <> "") Then strLocations = strLocations & ", " End If
        strLocations = strLocations & objState(i).Text
        If (i <> objLocations.length-1) Then strLocations = strLocations & "<br />" End If
    next
Else
    strLocations = ""
End If
Set objSite = xmldoc.selectNodes(strSelectNode & "site")
If objSite.length > 0 Then strSite = objSite(0).Text Else strSite = "" End If
Set objDescription = xmldoc.selectNodes(strSelectNode & "description")
If objDescription.length > 0 Then strDescription = objDescription(0).Text Else strDescription = "" End If

If objId.length > 0 Then
%>
<link title="combined" rel="stylesheet" type="text/css" media="screen" href="mapview.css" />
<script type="text/javascript" src="http://dev.virtualearth.net/mapcontrol/mapcontrol.ashx?v=6"></script>
<script language="javascript" type="text/javascript">
var map = null;
var mapZoom = 7;
var arrPin = new Array();
var pinid = 0;
var panningpin = null;
var panningBehavior = false;      
var site = "<%= strSite %>";
if (site == "Great Barrier Reef" || site == "Tsodilo")
{
    mapZoom = 5;
}	
function GetMap()
{
	map = new VEMap('myMap');
	var defaultMapLoc = GetDefaultMapLoc();
	map.SetDashboardSize(VEDashboardSize.Normal);	
	map.onLoadMap = MapLoaded;
	map.LoadMap(defaultMapLoc,mapZoom,'r');
	map.AttachEvent("onmousewheel", onscrollwheel);
	
	AddPushpin();
}

function onscrollwheel(e)
{
     return true;
} 

function GetDefaultMapLoc()
{
	var latLon = new VELatLong('<%=numLat%>','<%=numLong%>');
	return latLon;
}

function AddPushpin()
{
	var loc = new VELatLong('<%=numLat%>','<%=numLong%>');
	var customIcon = "<div class='pinImgOff' onmouseover='this.className=\"pinImgOn\";' onmouseout='this.className=\"pinImgOff\";'>1</div>";
	
	var shape = new VEShape(VEShapeType.Pushpin, loc);
	shape.SetCustomIcon(customIcon);
	map.AddShape(shape);
}

function MapLoaded(){}

function addEvent(obj,evType,fn)
{
	if(obj.addEventListener)
	{
		obj.addEventListener(evType,fn,false);
		return true;
	}
	else if(obj.attachEvent)
	{
		var r=obj.attachEvent("on"+evType,fn);
		return r;
	}
	else
	{
		return false;
	}
}

addEvent(window, 'load', GetMap);
</script>
<style type="text/css">
.pinImgOff {
    width:29px;
    height:32px;
    filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(
    src='http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_pinicon_off.png', sizingMethod='scale'); 
    color: #ffffff;
    text-align:center;
    font-size:11px;
    font-weight:normal;
    padding-top: 3px;
    cursor:pointer;
}
.pinImgOn {
    width:29px;
    height:32px;
    filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(
    src='http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_pinicon_on.png', sizingMethod='scale');
    color: #cc6600;
    text-align:center;
    font-size:11px;
    font-weight:normal;
    padding-top: 3px;
    cursor:pointer;
}
html>body .pinImgOff {
    background-image:url(http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_pinicon_off.gif);
    background-repeat:no-repeat;
    width:23px;
    height:26px;
    }
html>body .pinImgOn {
    background-image:url(http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_pinicon_on.gif);
    background-repeat:no-repeat;
    width:23px;
    height:26px;
    }
.gridIcon {
    background-image:url(http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_grid_button.gif);
    background-repeat:no-repeat;
    width:24px;
    height:20px;
    color: #ffffff;
    text-align:center;
    font-size:11px;
    font-weight:normal;
    padding-top: 3px;
}
.gridIconOff {
    background-image:url(http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_grid_button_selected.gif);
    background-repeat:no-repeat;
    width:24px;
    height:20px;
    color: #cc6600;
    text-align:center;
    font-size:11px;
    font-weight:normal;
    padding-top: 3px;
}
.MSVE_ZoomBar {
    background-color:#5a7a9b;
    filter:alpha(opacity=10);
    opacity:0.1;
}
</style>
<table cellpadding="0" cellspacing="0" border="0" width="400">
    <tr>
        <td align="left" valign="top" colspan="2" style="color:#264466; font-size:22px; font-family:Arial,helvetica, sans serif; font-weight:normal; padding-left:8px;"><%=strSite%></td>
    </tr>
    <tr>
        <td height="10" colspan="2"></td>
    </tr>
    <tr>
        <td width="400">
            <table cellpadding="0" cellspacing="0" border="0">
                <tr>
                    <td valign="top" width="95">
                        <div style="border:solid 1px #cccccc;">
                            <img id="unesco" src="http://whc.unesco.org/uploads/sites/site_<%=numId%>.jpg" alt="<%=strLocations%>" style="margin:1px 1px -2px 1px;" />
                        </div>                    
                    </td>
                    <td valign="top" style="font-family:Arial, helvetica, sans serif; font-size: 12px; color: #333333; padding-left:8px; line-height:18px;">
                        <strong>Region</strong>:<br /><%=strRegion%><br /><br /><strong>Location(s)</strong>:<br /><%=strLocations%><br /><br /><strong>Date of inscription</strong>:<br /><%=dtInscribeDate%>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td height="10" colspan="2"></td>
    </tr>
    <tr>
        <td colspan="2" style="font-family:Arial, helvetica, sans serif; font-size: 12px; color: #333333;padding-left:8px;"><%=strDescription%></td>
    </tr>
    <tr>
        <td height="20" colspan="2"></td>
    </tr>
    <tr>
        <td colspan="2" style="color:#264466; font-size:18px; font-family:Arial, helvetica, sans serif; font-weight:normal; padding-left:8px;">
            Explore the area nearby
        </td>
    </tr>
    <tr>
        <td height="10"></td>
    </tr>
    <tr>
        <td colspan="2" style="border:solid 1px #cccccc;">
            <div id='myMap' style="position:relative; width:400px; height:200px; margin:2px 2px 2px 2px;"></div>
        </td>
    </tr>
    <tr>
        <td height="20"></td>
    </tr>
    <tr>
        <td>
            <table cellpadding="0" cellspacing="0" border="0">
                <tr>
                    <td>
                        <table>
                            <tr>
                                <td style="font-family:Arial, helvetica, sans serif; font-size: 12px; color: #333333; padding-left:8px;" align="left">
                                    <div class="gridIcon" onmouseover="this.className='gridIconOff';"  onmouseout="this.className='gridIcon';" style="width: 24px; vertical-align: middle;">1</div>
                                </td>
                                <td style="font-family:Arial, helvetica, sans serif; font-size: 12px; color: #333333;">
                                    <strong><%=strSite%></strong>
                                </td>
                            </tr>
                         </table>
                       </td>
                     </tr>
                <tr>
                    <td style="font-family:Arial, helvetica, sans serif; font-size: 12px; padding-top:10px; padding-left:12px;">
                        <a href="find.asp" onmouseover="this.className='sitepageLinkOn'" onmouseout="this.className='sitepageLinkOff'" class="sitepageLink"><span style="font-size:12px;">Find another World Heritage site</span></a> <img src="http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_arrow-sm.gif" name="myArrow">
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<script language="javascript" type="text/javascript">
if (navigator.userAgent.indexOf("Firefox") != -1)
{
    document.getElementById("unesco").style.cssText = "margin:1px 1px 1px 1px;";
}
</script>
<%
Else  
    Response.Redirect("find.asp")    
End If 
%>