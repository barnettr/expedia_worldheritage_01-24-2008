var pinID = 1;
var selStyle = "r";  // a, h, r
var zoom = 3;
var selMode = VEMapMode.Mode2D;
var fixed = 0;
var showSwitch = 1

var map = null;
var arrPin = new Array();
var pinid = 0;
var panningpin = null;
var panningBehavior = false;
var nextZ = 1000;



function GetMap()
{
	map = new VEMap('myMap');
	var defaultMapLoc = GetDefaultMapLoc();
	map.SetDashboardSize(VEDashboardSize.Normal);	
	map.onLoadMap = MapLoaded;
	map.LoadMap(defaultMapLoc,3,'r');
	map.ClearInfoBoxStyles();
	map.AttachEvent("onendpan", OnMapMoved);
	map.AttachEvent("onchangeview", OnMapMoved);
	map.AttachEvent("onmouseover", mouseoverPin);
	map.AttachEvent("onmouseout", mouseoutPin);	
	getBrowser();
}

function getBrowser()
{
    if (navigator.userAgent.indexOf("MSIE") != -1)
    {
        g_isIE = true;
        AddPushpins();
        
        document.getElementById("logo").style.cssText = "display:none; z-index:0; visibility:hidden;";
        //document.getElementById("inprogress_img").style.cssText = "display:block; z-index:30; position:absolute; left:850px; top:280px; width:100px; height:20px; background-color:#d7e8e6; border:solid 2px #a31818; font-weight:bold; color:#666666; font-family:Arial; font-size:11px; vertical-align:middle;";
        if (navigator.userAgent.indexOf("MSIE 7") != -1)
        {
            g_isIE7 = true;
        }
    }
    else if (navigator.userAgent.indexOf("Firefox") != -1)
    {
        g_isFirefox = true;
        AddPushpins();
    }
}

function MapLoaded()
{
	//document.getElementById('preload').style.display='none';
	//document.getElementById('myGrid').style.display='block';
	document.getElementById('myGrid').style.visibility='visible';
	document.getElementById('showhide').style.cssText = "display:block; padding:1px 1px 1px 1px; margin:0px; border:1px solid #e6e7e7; font-weight:bold; background-color:#f4f9f9; color:#000000; text-decoration:none; font-family:Arial; font-size:8px; position:absolute; left:500px; top:167px; z-index:30;";
	document.getElementById("hide").style.cssText = "width:10px; height:10px; display:inline;";
	document.getElementById('expedia_logo').style.cssText = "display:block; position:absolute; top:476px; width:93px; height:29px; left:594px; z-index:30; filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(src='logo.png', sizingMethod='scale');";
	
	if (navigator.userAgent.indexOf("Firefox") != -1)
	{
	    document.getElementById('showhide').style.cssText = "display:block; padding:1px 1px 1px 1px; margin:0px; border:1px solid #e6e7e7; font-weight:bold; background-color:#f4f9f9; color:#000000; text-decoration:none; font-family:Arial; font-size:8px; position:absolute; left:497px; top:159px; z-index:30;";
	    document.getElementById('expedia_logo').style.cssText = "display:block; position:absolute; top:470px; width:92px; height:29px; left:588px; z-index:30; filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(src='logo.png', sizingMethod='scale');";
	}
	
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
		var pin = BuildPushpin(arrPin[x].id,arrPin[x].site,arrPin[x].suitcase,arrPin[x].flag,arrPin[x].desc,arrPin[x].img,arrPin[x].lat,arrPin[x].lng)
		map.AddShape(pin);
		arrPin[x].pinid = pin.GetID();
	}
}

function BuildPushpin(id,site,suitcase,flag,desc,img,lat,lng)
{ 
	var loc = new VELatLong(lat,lng);
	var title = "<table border='0' cellpadding='0' cellspacing='0' style='width:100%'><tr><td align='left' valign='top' style='width:14px;'><span style='padding: 1px 3px 1px 3px; margin: 2px; border: 1px solid #EEE; font-weight: bold; background-color: #4985c7; color: #FFF;'>" + id + "</span></td><td align='left' valign='top' style='text-align:left; padding-left:6px;'><a href='' style='text-decoration:none; border:0px; font-weight:bold; text-align:left; color:#4985c7'>" + site + "</a></td><td align='right' valign='top' style='width:41px; height:34px; padding-left:4px;'><img src='" + suitcase + "' style='border:none;'></td></tr></table>"
	var description = "<img src='" + flag + "' style='border: solid 1px #000000; margin-bottom:4px;'><br /><div id='PushpinBalloonContainer' style='margin-left:0;'><table border='0' cellpadding='0' cellspacing='0'><tr><td align='left' valign='top' style='font-family: arial; font-size: 1.0em; color: #000000; padding-right:6px; text-align:justify;'><strong style='font-family:arial; font-size:11px;'>Brief Description</strong><br />" + desc + "</td><td align='right' valign='top' style='padding-right:4px;'><img src='" + img + "' style='border: solid 1px #979797; margin-top:16px;' align='right'></td></tr></table></div>";
	var customIcon = "<div style='padding: 1px 1px 1px 1px; margin: 0px; border: 1px solid #EEE; font-weight:bold; background-color: #b0e3fa; color: #00156e; text-decoration:none; font-family:verdana; font-size:11px;'>" + id + "</div>";
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
		var pin = map.GetShapeByID(e.elementID);
		pin.SetZIndex(getHighestZIndex());
		var row = 'row_'+getRowID(map.GetShapeByID(e.elementID).GetID());
		showRow(row,'pin');
	}
}

function mouseoutPin(e)
{
	if (e.elementID != null)
    {
		//var pin = map.GetShapeByID(e.elementID);
		//pin.SetZIndex(1000);
		var row = 'row_'+getRowID(map.GetShapeByID(e.elementID).GetID());
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
			map.ShowInfoBox(pin);
		}
		else
		{
			panningpin = pin;
			CenterPin();
		}
	}
}

function GridHideInfoBox(obj)
{
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
		setTimeout("map.ShowInfoBox(panningpin)",10);
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
	//var rowObj = document.getElementById(id);
	if(type == 'pin')
	{
		document.getElementById(id).scrollIntoView(true);
	}
	document.getElementById(id).className='rowHover';
}

function showRow(id,type)
{
	//var rowObj = document.getElementById(id);
	if(type == 'pin')
	{
		document.getElementById(id).scrollIntoView(true);
	}
	document.getElementById(id).className='rowHover';
}

function hideRow(id)
{
	document.getElementById(id).className='siteRow';
}

function getHighestZIndex()
{
	//for (x = 0; x < arrPin.length; x++)
	//{
	//	var pin = map.GetShapeByID(arrPin[x].pinid);
	//	alert(pin.GetID()+" - "+pin.GetZIndex());
	//}
	return nextZ += 1;
}