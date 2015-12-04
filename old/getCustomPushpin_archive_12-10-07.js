function getCustomPushpin()
{
    var xmlDoc = "";
    var xmlObj = "";
    var clickedShape = null;
    
    var Pushpin_Group = "";
    var Pushpin_Region = "";
    var sFile = "";
    
    var currentPage = "";
    var pageSize = "";
    var sortCol = "";
    var sortDir = "";
    
    Pushpin_Region = document.getElementById("sRegion").value;
    Pushpin_State = document.getElementById("sState").value;
    sFile = document.getElementById("sFile").value;
	
	currentPage = document.getElementById("paramCurrentPage").value;
	pageSize = document.getElementById("paramPageSize").value;
	sortCol = document.getElementById("paramSortCol").value;
	sortDir = document.getElementById("paramSortDir").value;
	sortType = document.getElementById("paramSortType").value;
	
	
    if (Pushpin_Region != "" && Pushpin_State == "")  // Initial dropdown box: Region
    {
        switch (Pushpin_Region)  // Five initial global location latitude's & longitude's
        {
            case "Europe":
            var latLon = new VELatLong(46.485070, 14.5097);
            break;
            case "Africa":
            var latLon = new VELatLong(16.98485, 0.213671);
            break;
            case "Arab States":
            var latLon = new VELatLong(23.530939, 45.044498);
            break;
            case "Asia":
            var latLon = new VELatLong(39.84555556, 116.44472220);
            break;
            case "Latin America":
            var latLon = new VELatLong(3.96666667, -81.61666667);
            break;
            case "North America":
            var latLon = new VELatLong(45.291157, -93.265892);
            break;
            case "Caribbean":
            var latLon = new VELatLong(18.48333333, -69.91666667);
            break;
            case "Pacific":
            var latLon = new VELatLong(-8.54333000, 119.48944000);
            break;
        }
        
        zoom = 3;
        map = new VEMap('myMap');
        map.SetDashboardSize(VEDashboardSize.Normal);
        map.LoadMap(latLon, zoom, selStyle, fixed, selMode, showSwitch);
        map.SetScaleBarDistanceUnit(VEDistanceUnit.Miles);
        
		// This works but is commented out to try to use xslt
        //xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
        //xmlDoc.async = false;
        //xmlDoc.load(sFile);
		
		// Load XML and apply XSLT
		var xslt = new ActiveXObject("Msxml2.XSLTemplate");
		var xslDoc = new ActiveXObject("Msxml2.FreeThreadedDOMDocument");
		var xslProc;
		xslDoc.async = false;
		xslDoc.resolveExternals = false;
		xslDoc.load("mapSorted.xslt");
		xslt.stylesheet = xslDoc;
		var xmlDoc = new ActiveXObject("Msxml2.DOMDocument");
		xmlDoc.async = false;
		xmlDoc.resolveExternals = false;
		xmlDoc.load(sFile);
		xslProc = xslt.createProcessor();
		xslProc.input = xmlDoc;
	
		xslProc.addParameter("SortCol", sortCol);
		xslProc.addParameter("SortDir", sortDir);
		xslProc.addParameter("SortType", sortType);
	
		xslProc.transform();
		var xmlString;
		xmlString = xslProc.output;
		var xmlSorted = new ActiveXObject("Msxml2.DOMDocument");
		xmlSorted.loadXML(xmlString);
        xmlObj = xmlSorted.documentElement; 
        //alert(xmlObj.nodeName);
		
		//xmlObj = xmlDoc.documentElement; 
        //alert(xmlObj.nodeName);
        
        var b = 1;
        
        for (var a = 0; a < xmlObj.childNodes.length; a++)
        {
            if (xmlObj.childNodes[a].childNodes[3].tagName == "region" && xmlObj.childNodes[a].childNodes[3].text == Pushpin_Region)
            {
                if ((a >= 0 + (currentPage -1) * pageSize) && (a < (0 + currentPage * pageSize)))
                { 
                    var flags = (xmlObj.childNodes[a].childNodes[2].tagName == "code") ? xmlObj.childNodes[a].childNodes[2].text.split(",") : "";
                    var location = (xmlObj.childNodes[a].childNodes[4].tagName == "location") ? xmlObj.childNodes[a].childNodes[4].text : "";
                    var suitcase = "<img src='suitcase.gif' style='border:none;'";
                    var lat = (xmlObj.childNodes[a].tagName == "row") ? parseFloat(xmlObj.childNodes[a].getAttribute("lat")) : "";
                    var lng = (xmlObj.childNodes[a].tagName == "row") ? parseFloat(xmlObj.childNodes[a].getAttribute("long")) : "";
                    var id = (xmlObj.childNodes[a].tagName == "row") ? xmlObj.childNodes[a].getAttribute("id") : "";
                    var image = (xmlObj.childNodes[a].tagName == "row") ? "http://whc.unesco.org/uploads/sites/site_" + xmlObj.childNodes[a].getAttribute("id") + ".jpg" : "";
                    var description = (xmlObj.childNodes[a].childNodes[7].tagName == "description") ? xmlObj.childNodes[a].childNodes[7].text : "";
                    var points = new VELatLong(lat,lng);
                    
                    var infobox = "<p>";
                    for (x = 0; x < flags.length; x++ )
                    {
                        infobox += (flags[x]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[x] + ".gif' style='border: solid 1px #000000;'> " : "";
                    }
                    infobox += "</p><div id='PushpinBalloonContainer' style='margin-left:0;'><table border='0' cellpadding='0' cellspacing='0'><tr><td align='left' valign='top' style='font-family: arial; font-size: 1.0em; color: #000000; padding-right:6px; text-align:justify;'><strong style='font-family:arial; font-size:11px;'>Brief Description</strong><br />" + description + "</td><td align='right' valign='top' style='padding-right:4px;'><img src='" + image + "' style='border: solid 1px #979797;' align='right'></td></tr></table></div>";
                    
                    if (currentPage == 1)
                    {
                        var title = (xmlObj.childNodes[a].childNodes[6].tagName == "site") ? "<table id='tpin_"+id+"' border='0' cellpadding='0' cellspacing='0' style='width:100%'><tr><td align='left' valign='top' style='width:14px;'><span class='counter'>" + b + "</span></td><td align='left' valign='top' style='text-align:left; padding-left:4px;'><a href='' style='text-decoration:none; border:0px; font-weight:bold; text-align:left; color:#4985c7'>" + xmlObj.childNodes[a].childNodes[6].text + "</a></td><td align='right' valign='top' style='width:41px; height:34px; padding-left:4px;'>" + suitcase + "</td></tr></table>" : "";
                        var customIcon = "<div id='pin_"+id+"' style='padding: 1px 1px 1px 1px; margin: 0px; border: 1px solid #EEE; font-weight:bold; background-color: #b0e3fa; color: #00156e; text-decoration:none; font-family:verdana; font-size:11px;'>" + b + "</div>";
                    }
                    else
                    {
                        var title = (xmlObj.childNodes[a].childNodes[6].tagName == "site") ? "<table id='tpin_"+id+"' border='0' cellpadding='0' cellspacing='0' style='width:100%'><tr><td align='left' valign='top' style='width:14px;'><span class='counter'>" + (b+((currentPage-1) * pageSize)) + "</span></td><td align='left' valign='top' style='text-align:left; padding-left:4px;'><a href='' style='text-decoration:none; border:0px; font-weight:bold; text-align:left; color:#4985c7'>" + xmlObj.childNodes[a].childNodes[6].text + "</a></td><td align='right' valign='top' style='width:41px; height:34px; padding-left:4px;'>" + suitcase + "</td></tr></table>" : "";
                        var customIcon = "<div id='pin_"+id+"' style='padding: 1px 1px 1px 1px; margin: 0px; border: 1px solid #EEE; font-weight:bold; background-color: #b0e3fa; color: #00156e; text-decoration:none; font-family:verdana; font-size:11px;'>" + (b+((currentPage-1) * pageSize)) + "</div>";
                    }   
                      b++;
                    
                    var shape = new VEShape(VEShapeType.Pushpin, points);
                    shape.SetCustomIcon(customIcon);
                    shape.SetTitle(title);
                    shape.SetDescription(infobox);
                    map.ClearInfoBoxStyles();
                    map.AttachEvent("onmouseover", OnMouseOver);
                    map.AttachEvent("onmouseout", OnMouseOut);
                    map.AddShape(shape);
                    
                    function OnMouseOver(e)
                    {
                        if (e.elementID)
                        {
                            y = window.event.y;
                            x = window.event.x;
                            var pixel = new VEPixel((x), (y+50));
                            var latLong = map.PixelToLatLong(pixel);
                            var hoveredShape = map.GetShapeByID(e.elementID);

                            if (hoveredShape)
                            {
                                //map.SetCenter(hoveredShape.GetIconAnchor());
                                map.ShowInfoBox(hoveredShape, latLong);
                                //CenterPin();
                                return true;
                            }   
                        }
                    }
                    
                    function OnMouseOut(e)
                    {
                        if (e.elementID)
                        {
                            VEHideVEShapeERO(false);
                        }
                    }
                }  
            }
        }
        setTimeout('document.getElementById("inprogress_img").style.cssText = "display:none;"', 2000);
        setTimeout('document.getElementById("loader").style.cssText = "display:none;"', 2000);   
    }
    else if (Pushpin_Region != "" && Pushpin_State != "")  // Two dropdown boxes: Region & State
    {
        var xslt = new ActiveXObject("Msxml2.XSLTemplate");
		var xslDoc = new ActiveXObject("Msxml2.FreeThreadedDOMDocument");
		var xslProc;
		xslDoc.async = false;
		xslDoc.resolveExternals = false;
		xslDoc.load("mapSorted.xslt");
		xslt.stylesheet = xslDoc;
		var xmlDoc = new ActiveXObject("Msxml2.DOMDocument");
		xmlDoc.async = false;
		xmlDoc.resolveExternals = false;
		xmlDoc.load(sFile);
		xslProc = xslt.createProcessor();
		xslProc.input = xmlDoc;
	
		xslProc.addParameter("SortCol", sortCol);
		xslProc.addParameter("SortDir", sortDir);
		xslProc.addParameter("SortType", sortType);
	
		xslProc.transform();
		var xmlString;
		xmlString = xslProc.output;
		var xmlSorted = new ActiveXObject("Msxml2.DOMDocument");
		xmlSorted.loadXML(xmlString);
        xmlObj = xmlSorted.documentElement;
        var b = 1;
        
        for (var x = 0; x < xmlObj.childNodes.length; x++)
        {
            if (xmlObj.childNodes[x].tagName == "row" && xmlObj.childNodes[x].childNodes[3].text == Pushpin_Region && xmlObj.childNodes[x].childNodes[5].text == Pushpin_State)
            {
                var lat = parseFloat(xmlObj.childNodes[x].getAttribute("lat"));
                var lng = parseFloat(xmlObj.childNodes[x].getAttribute("long"));
                var latLon = new VELatLong(lat,lng);
                var zoom = 5;
            }
        }
        
        map = new VEMap('myMap');
        map.LoadMap(latLon, zoom, selStyle, fixed, selMode, showSwitch);
        map.SetScaleBarDistanceUnit(VEDistanceUnit.Miles);
       
   
        for (var a = 0; a < xmlObj.childNodes.length; a++)
        {
            if (xmlObj.childNodes[a].tagName == "row" && xmlObj.childNodes[a].childNodes[3].text == Pushpin_Region && xmlObj.childNodes[a].childNodes[5].text == Pushpin_State)
            {
                //if ((a >= 0 + (currentPage -1) * pageSize) && (a < (0 + currentPage * pageSize)))
                //{
                    var flags = (xmlObj.childNodes[a].childNodes[2].tagName == "code") ? xmlObj.childNodes[a].childNodes[2].text.split(",") : "";
                    var location = (xmlObj.childNodes[a].childNodes[4].tagName == "location") ? xmlObj.childNodes[a].childNodes[4].text : "";
                    var suitcase = "<img src='suitcase.gif' style='border:none;'";
                    //var title = (xmlObj.childNodes[a].childNodes[6].tagName == "site") ? "<table border='0' cellpadding='0' cellspacing='0' style='width:100%'><tr><td align='left' valign='top' style='width:14px;'><span class='counter'>" + b + "</span></td><td align='left' valign='top' style='text-align:left;'><a href='' style='text-decoration:none; border:0px; font-weight:bold; text-align:left;'>" + xmlObj.childNodes[a].childNodes[6].text + "</a></td><td align='right' valign='top' style='width:41px; height:34px; padding-left:4px;'>" + suitcase + "</td></tr></table>" : "";
                    var lat = (xmlObj.childNodes[a].tagName == "row") ? parseFloat(xmlObj.childNodes[a].getAttribute("lat")) : "";
                    var lng = (xmlObj.childNodes[a].tagName == "row") ? parseFloat(xmlObj.childNodes[a].getAttribute("long")) : "";
                    var image = (xmlObj.childNodes[a].tagName == "row") ? "http://whc.unesco.org/uploads/sites/site_" + xmlObj.childNodes[a].getAttribute("id") + ".jpg" : "";
                    var description = (xmlObj.childNodes[a].childNodes[7].tagName == "description") ? xmlObj.childNodes[a].childNodes[7].text : "";
                    var points = new VELatLong(lat,lng);
                    
                    var infobox = "<p>";
                    for (x = 0; x < flags.length; x++ )
                    {
                        infobox += (flags[x]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[x] + ".gif' style='border: solid 1px #000000;'> " : "";
                    }
                    infobox += "</p><div id='PushpinBalloonContainer' style='margin-left:0;'><table border='0' cellpadding='0' cellspacing='0'><tr><td align='left' valign='top' style='font-family: arial; font-size: 1.0em; color: #000000; padding-right:6px; text-align:justify;'><strong style='font-family:arial; font-size:11px;'>Brief Description</strong><br />" + description + "</td><td align='right' valign='top' style='padding-right:2px;'><a href=''><img src='" + image + "' style='border: solid 1px #979797;' align='right'></a></td></tr></table></div>";
                    
                    if (currentPage == 1)
                    {
                        var title = (xmlObj.childNodes[a].childNodes[6].tagName == "site") ? "<table id='tpin_"+id+"' border='0' cellpadding='0' cellspacing='0' style='width:100%'><tr><td align='left' valign='top' style='width:14px;'><span class='counter'>" + b + "</span></td><td align='left' valign='top' style='text-align:left; padding-left:4px;'><a href='' style='text-decoration:none; border:0px; font-weight:bold; text-align:left; color:#4985c7'>" + xmlObj.childNodes[a].childNodes[6].text + "</a></td><td align='right' valign='top' style='width:41px; height:34px; padding-left:4px;'>" + suitcase + "</td></tr></table>" : "";
                        var customIcon = "<div id='pin_"+id+"' style='padding: 1px 1px 1px 1px; margin: 0px; border: 1px solid #EEE; font-weight:bold; background-color: #b0e3fa; color: #00156e; text-decoration:none; font-family:verdana; font-size:11px;'>" + b + "</div>";
                    }
                    else
                    {
                        var title = (xmlObj.childNodes[a].childNodes[6].tagName == "site") ? "<table id='tpin_"+id+"' border='0' cellpadding='0' cellspacing='0' style='width:100%'><tr><td align='left' valign='top' style='width:14px;'><span class='counter'>" + (b+((currentPage-1) * pageSize)) + "</span></td><td align='left' valign='top' style='text-align:left; padding-left:4px;'><a href='' style='text-decoration:none; border:0px; font-weight:bold; text-align:left; color:#4985c7'>" + xmlObj.childNodes[a].childNodes[6].text + "</a></td><td align='right' valign='top' style='width:41px; height:34px; padding-left:4px;'>" + suitcase + "</td></tr></table>" : "";
                        var customIcon = "<div id='pin_"+id+"' style='padding: 1px 1px 1px 1px; margin: 0px; border: 1px solid #EEE; font-weight:bold; background-color: #b0e3fa; color: #00156e; text-decoration:none; font-family:verdana; font-size:11px;'>" + (b+((currentPage-1) * pageSize)) + "</div>";
                    }
                      b++;
                    
                    var shape = new VEShape(VEShapeType.Pushpin, points);
                    shape.SetCustomIcon(customIcon);
                    shape.SetTitle(title);
                    shape.SetDescription(infobox);
                    map.AttachEvent("onmouseover", OnMouseOver);
                    map.AttachEvent("onmouseout", OnMouseOut);
                    map.ClearInfoBoxStyles();
                    map.AddShape(shape);
                    
                    function OnMouseOver(e)
                    {
                        if (e.elementID)
                        {
                            y = window.event.y;
                            x = window.event.x;
                            var pixel = new VEPixel((x), (y+50));
                            var latLong = map.PixelToLatLong(pixel);
                            var hoveredShape = map.GetShapeByID(e.elementID);

                            if (hoveredShape)
                            {
                                //map.SetCenter(hoveredShape.GetIconAnchor());
                                map.ShowInfoBox(hoveredShape, latLong);
                                return true;
                            }   
                        }
                    }
                    
                    function OnMouseOut(e)
                    {
                        if (e.elementID)
                        {
                            VEHideVEShapeERO(false);
                        }
                    }
                //}    
            }
        }
        setTimeout('document.getElementById("inprogress_img").style.cssText = "display:none;"', 2000);
        setTimeout('document.getElementById("loader").style.cssText = "display:none;"', 2000);
    }
}
