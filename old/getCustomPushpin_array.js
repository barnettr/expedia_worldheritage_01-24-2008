function getCustomPushpin()
{
    var pinID = 1;
    var selStyle = "r";  // a, h, r
    var zoom = 3;
    var selMode = VEMapMode.Mode2D;
    var fixed = 0;
    var showSwitch = 1
    var map = null;
    
    var xmlDoc = "";
    var xmlObj = "";
    
    var Pushpin_Group = "";
    var Pushpin_Region = "";
    var sFile = "";
    
    var currentPage = "";
    var pageSize = "";
    var sortCol = "";
    var sortDir = "";
    
    Pushpin_Group = document.getElementById("sGroup").value;
    Pushpin_Region = document.getElementById("sRegion").value;
    sFile = document.getElementById("sFile").value;
	
	currentPage = document.getElementById("paramCurrentPage").value;
	pageSize = document.getElementById("paramPageSize").value;
	sortCol = document.getElementById("paramSortCol").value;
	sortDir = document.getElementById("paramSortDir").value;
	//alert(currentPage+" "+pageSize+" "+sortCol+" "+sortDir);
	
	function OnMouseOver(e)
    {
        return true;
    }

    function OnClick(e)
    {
        if (e.elementID)
        {
            var clickedShape = map.GetShapeByID(e.elementID);
            
            if (clickedShape)
                map.ShowInfoBox(clickedShape);
        }
    }
       
    if (Pushpin_Group != "" && Pushpin_Region == "")  // Initial dropdown box: Region
    {
        switch (Pushpin_Group)  // Five initial global location latitude's & longitude's
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
        
        map = new VEMap('myMap');
        map.LoadMap(latLon, zoom, selStyle, fixed, selMode, showSwitch);
        map.AttachEvent("onmouseover", OnMouseOver);
        map.AttachEvent("onclick", OnClick);
        map.SetScaleBarDistanceUnit(VEDistanceUnit.Miles);
        map.ClearInfoBoxStyles();
        
        
        var allPushpins = new Array(
        xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
        xmlDoc.async = false;
        xmlDoc.load(sFile);
        xmlObj = xmlDoc.documentElement;
        
        var b = 1;
        
        for (var a = 0; a < xmlObj.childNodes.length; a++)
        {
            if (xmlObj.childNodes[a].childNodes[3].tagName == "region" && xmlObj.childNodes[a].childNodes[3].text == Pushpin_Group)
            {
                if ((a >= 1 + (currentPage -1) * pageSize) && (a < (1 + currentPage * pageSize)))
                {
                    var flags = (xmlObj.childNodes[a].childNodes[2].tagName == "code") ? xmlObj.childNodes[a].childNodes[2].text.split(",") : "";
                    var location = (xmlObj.childNodes[a].childNodes[4].tagName == "location") ? xmlObj.childNodes[a].childNodes[4].text : "";
                    var suitcase = "<img src='suitcase.gif' style='border:none;'";
                    var title = (xmlObj.childNodes[a].childNodes[6].tagName == "site") ? "<table border='0' cellpadding='0' cellspacing='0' style='width:100%'><tr><td align='left' valign='top' style='width:14px;'><span class='counter'>" + b + "</span></td><td align='left' valign='top' style='text-align:left; padding-left:4px;'><a href='' style='text-decoration:none; border:0px; font-weight:bold; text-align:left; color:#4985c7'>" + xmlObj.childNodes[a].childNodes[6].text + "</a></td><td align='right' valign='top' style='width:41px; height:34px; padding-left:4px;'>" + suitcase + "</td></tr></table>" : "";
                    var lat = (xmlObj.childNodes[a].tagName == "row") ? parseFloat(xmlObj.childNodes[a].getAttribute("lat")) : "";
                    var lng = (xmlObj.childNodes[a].tagName == "row") ? parseFloat(xmlObj.childNodes[a].getAttribute("long")) : "";
                    var image = (xmlObj.childNodes[a].tagName == "row") ? "http://whc.unesco.org/uploads/sites/site_" + xmlObj.childNodes[a].getAttribute("id") + ".jpg" : "";
                    var description = (xmlObj.childNodes[a].childNodes[7].tagName == "description") ? xmlObj.childNodes[a].childNodes[7].text : "";
                    var points = new VELatLong(lat,lng);
                    
                    var infobox = "<p><img src='http://whc.unesco.org/uploads/states/small/" + flags[0] + ".gif' style='border: solid 1px #000000;'> ";
                    infobox += (flags[1]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[1] + ".gif' style='border: solid 1px #000000;'> " : "";
                    infobox += (flags[2]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[2] + ".gif' style='border: solid 1px #000000;'> " : "";
                    infobox += (flags[3]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[3] + ".gif' style='border: solid 1px #000000;'> " : "";
                    infobox += (flags[4]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[4] + ".gif' style='border: solid 1px #000000;'> " : "";
                    infobox += (flags[5]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[5] + ".gif' style='border: solid 1px #000000;'> " : "";
                    infobox += (flags[6]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[6] + ".gif' style='border: solid 1px #000000;'> " : "";
                    infobox += (flags[7]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[7] + ".gif' style='border: solid 1px #000000;'> " : "";
                    infobox += (flags[8]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[8] + ".gif' style='border: solid 1px #000000;'> " : "";
                    infobox += (flags[9]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[9] + ".gif' style='border: solid 1px #000000;'> " : "";
                    infobox += (flags[10]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[10] + ".gif' style='border: solid 1px #000000;'> " : "";
                    infobox += (flags[11]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[11] + ".gif' style='border: solid 1px #000000;'> " : "";
                    infobox += "</p><div id='PushpinBalloonContainer' style='margin-left:0;'><table border='0' cellpadding='0' cellspacing='0'><tr><td align='left' valign='top' style='font-family: arial; font-size: 1.0em; color: #000000; padding-right:6px; text-align:justify;'><strong style='font-family:arial; font-size:11px;'>Brief Description</strong><br />" + description + "</td><td align='right' valign='top' style='padding-right:2px;'><a href=''><img src='" + image + "' style='border: solid 1px #979797;' align='right'></a></td></tr></table></div>";
                    
                    var customIcon = "<div style='padding: 1px 1px 1px 1px; margin: 0px; border: 1px solid #EEE; font-weight:bold; background-color: #b0e3fa; color: #00156e; text-decoration:none; font-family:verdana; font-size:11px;'>" + b + "</div>";
                      b++;
                    
                    var shape = new VEShape(VEShapeType.Pushpin, points);
                    alert(typeof shape);
                    shape.SetCustomIcon(customIcon);
                    shape.SetTitle(title);
                    shape.SetDescription(infobox);
                    map.ClearInfoBoxStyles();
                    map.AddShape(shape);
                }
                //var pin = map.AddPushpin(points);
                //pin.SetCustomIcon(customIcon);
                //pin.SetTitle(title);
                //pin.SetDescription(infobox);    
            }
        }
        )
        setTimeout('document.getElementById("inprogress_img").style.cssText = "display:none;"', 2000);
        setTimeout('document.getElementById("loader").style.cssText = "display:none;"', 2000);   
    }
    else if (Pushpin_Group != "" && Pushpin_Region != "")  // Two dropdown boxes: Region & State
    {
        xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
        xmlDoc.async = false;
	    xmlDoc.load(sFile);
	    xmlObj = xmlDoc.documentElement;
        var b = 1;
        
        for (var x = 0; x < xmlObj.childNodes.length; x++)
        {
            if (xmlObj.childNodes[x].tagName == "row" && xmlObj.childNodes[x].childNodes[3].text == Pushpin_Group && xmlObj.childNodes[x].childNodes[5].text == Pushpin_Region)
            {
                var lat = parseFloat(xmlObj.childNodes[x].getAttribute("lat"));
                var lng = parseFloat(xmlObj.childNodes[x].getAttribute("long"));
                var latLon = new VELatLong(lat,lng);
                var zoom = 5;
            }
        }
        
        map = new VEMap('myMap');
        map.LoadMap(latLon, zoom, selStyle, fixed, selMode, showSwitch);
        map.AttachEvent("onmouseover", OnMouseOver);
        map.AttachEvent("onclick", OnClick);
        map.SetScaleBarDistanceUnit(VEDistanceUnit.Miles);
        map.ClearInfoBoxStyles();
   
        for (var a = 0; a < xmlObj.childNodes.length; a++)
        {
            if (xmlObj.childNodes[a].tagName == "row" && xmlObj.childNodes[a].childNodes[3].text == Pushpin_Group && xmlObj.childNodes[a].childNodes[5].text == Pushpin_Region)
            {
                var flags = (xmlObj.childNodes[a].childNodes[2].tagName == "code") ? xmlObj.childNodes[a].childNodes[2].text.split(",") : "";
                var location = (xmlObj.childNodes[a].childNodes[4].tagName == "location") ? xmlObj.childNodes[a].childNodes[4].text : "";
                var suitcase = "<img src='suitcase.gif' style='border:none;'";
                var title = (xmlObj.childNodes[a].childNodes[6].tagName == "site") ? "<table border='0' cellpadding='0' cellspacing='0' style='width:100%'><tr><td align='left' valign='top' style='width:14px;'><span class='counter'>" + b + "</span></td><td align='left' valign='top' style='text-align:left;'><a href='' style='text-decoration:none; border:0px; font-weight:bold; text-align:left;'>" + xmlObj.childNodes[a].childNodes[6].text + "</a></td><td align='right' valign='top' style='width:41px; height:34px; padding-left:4px;'>" + suitcase + "</td></tr></table>" : "";
                var lat = (xmlObj.childNodes[a].tagName == "row") ? parseFloat(xmlObj.childNodes[a].getAttribute("lat")) : "";
                var lng = (xmlObj.childNodes[a].tagName == "row") ? parseFloat(xmlObj.childNodes[a].getAttribute("long")) : "";
                var image = (xmlObj.childNodes[a].tagName == "row") ? "http://whc.unesco.org/uploads/sites/site_" + xmlObj.childNodes[a].getAttribute("id") + ".jpg" : "";
                var description = (xmlObj.childNodes[a].childNodes[7].tagName == "description") ? xmlObj.childNodes[a].childNodes[7].text : "";
                var points = new VELatLong(lat,lng);
                
                var infobox = "<p><img src='http://whc.unesco.org/uploads/states/small/" + flags[0] + ".gif' style='border: solid 1px #000000;'> ";
                infobox += (flags[1]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[1] + ".gif' style='border: solid 1px #000000;'> " : "";
                infobox += (flags[2]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[2] + ".gif' style='border: solid 1px #000000;'> " : "";
                infobox += (flags[3]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[3] + ".gif' style='border: solid 1px #000000;'> " : "";
                infobox += (flags[4]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[4] + ".gif' style='border: solid 1px #000000;'> " : "";
                infobox += (flags[5]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[5] + ".gif' style='border: solid 1px #000000;'> " : "";
                infobox += (flags[6]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[6] + ".gif' style='border: solid 1px #000000;'> " : "";
                infobox += (flags[7]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[7] + ".gif' style='border: solid 1px #000000;'> " : "";
                infobox += (flags[8]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[8] + ".gif' style='border: solid 1px #000000;'> " : "";
                infobox += (flags[9]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[9] + ".gif' style='border: solid 1px #000000;'> " : "";
                infobox += (flags[10]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[10] + ".gif' style='border: solid 1px #000000;'> " : "";
                infobox += (flags[11]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[11] + ".gif' style='border: solid 1px #000000;'> " : "";
                infobox += "</p><div id='PushpinBalloonContainer' style='margin-left:0;'><table border='0' cellpadding='0' cellspacing='0'><tr><td align='left' valign='top' style='font-family: arial; font-size: 0.7em; color: #000000; padding-right:6px; text-align:justify;'><strong style='font-family:arial;'>Brief Description</strong><br />" + description + "</td><td align='right' valign='top' style='padding-right:2px;'><img src='" + image + "' style='border: solid 1px #979797;' align='right'></td></tr></table></div>";
                
                var customIcon = "<div style='padding: 1px 1px 1px 1px; margin: 0px; border: 1px solid #EEE; font-weight:bold; background-color: #b0e3fa; color: #00156e; text-decoration:none; font-family:verdana; font-size:11px;'>" + b + "</div>";
                  b++;
                
                var shape = new VEShape(VEShapeType.Pushpin, points);
                shape.SetCustomIcon(customIcon);
                shape.SetTitle(title);
                shape.SetDescription(infobox);
                map.ClearInfoBoxStyles();
                map.AddShape(shape);
//                var pin = map.AddPushpin(points);
//                pin.SetCustomIcon(customIcon);
//                pin.SetTitle(title);
//                pin.SetDescription(infobox);     
            }
        }
        setTimeout('document.getElementById("inprogress_img").style.cssText = "display:none;"', 2000);
        setTimeout('document.getElementById("loader").style.cssText = "display:none;"', 2000);
    }
//    else if (Pushpin_Group != "" && Pushpin_Region != "" && Pushpin_Location != "") // All three dropdown boxes selected by user: Region & State & Location
//    {
//        xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
//        xmlDoc.async = false;
//	    xmlDoc.load("worldheritage_" + Pushpin_Group + ".xml");
//	    xmlObj = xmlDoc.documentElement;
//        var b = 1;
//        
//        for (var x = 0; x < xmlObj.childNodes.length; x++)
//        {
//            if (xmlObj.childNodes[x].tagName == "row" && xmlObj.childNodes[x].childNodes[3].text == Pushpin_Group && xmlObj.childNodes[x].childNodes[5].text == Pushpin_Region && xmlObj.childNodes[x].childNodes[4].text == Pushpin_Location)
//            {
//                var lat = parseFloat(xmlObj.childNodes[x].getAttribute("lat"));
//                var lng = parseFloat(xmlObj.childNodes[x].getAttribute("long"));
//                var latLon = new VELatLong(lat,lng);
//                var zoom = 10;
//            }
//        }
//        
//        map = new VEMap('myMap');
//        map.LoadMap(latLon, zoom, selStyle, fixed, selMode, showSwitch);
//        map.AttachEvent("onmouseover", OnMouseOver);
//        map.AttachEvent("onclick", OnClick);
//        map.SetScaleBarDistanceUnit(VEDistanceUnit.Miles);
//        map.ClearInfoBoxStyles();
//        
//        for (var a = 0; a < xmlObj.childNodes.length; a++)
//        {
//            if (xmlObj.childNodes[a].tagName == "row" && xmlObj.childNodes[a].childNodes[3].text == Pushpin_Group && xmlObj.childNodes[a].childNodes[5].text == Pushpin_Region && xmlObj.childNodes[a].childNodes[4].text == Pushpin_Location)
//            {
//                var flags = (xmlObj.childNodes[a].childNodes[2].tagName == "code") ? xmlObj.childNodes[a].childNodes[2].text.split(",") : "";
//                var location = (xmlObj.childNodes[a].childNodes[4].tagName == "location") ? xmlObj.childNodes[a].childNodes[4].text : "";
//                var suitcase = "<img src='suitcase.gif' style='border:none;'";
//                var title = (xmlObj.childNodes[a].childNodes[6].tagName == "site") ? "<table border='0' cellpadding='0' cellspacing='0' style='width:100%'><tr><td align='left' valign='top' style='width:14px;'><span class='counter'>" + b + "</span></td><td align='left' valign='top' style='text-align:left;'><a href='' style='text-decoration:none; border:0px; font-weight:bold; text-align:left;'>" + xmlObj.childNodes[a].childNodes[6].text + "</a></td><td align='right' valign='top' style='width:41px; height:34px; padding-left:4px;'>" + suitcase + "</td></tr></table>" : "";
//                var lat = (xmlObj.childNodes[a].tagName == "row") ? parseFloat(xmlObj.childNodes[a].getAttribute("lat")) : "";
//                var lng = (xmlObj.childNodes[a].tagName == "row") ? parseFloat(xmlObj.childNodes[a].getAttribute("long")) : "";
//                var image = (xmlObj.childNodes[a].tagName == "row") ? "http://whc.unesco.org/uploads/sites/site_" + xmlObj.childNodes[a].getAttribute("id") + ".jpg" : "";
//                var description = (xmlObj.childNodes[a].childNodes[7].tagName == "description") ? xmlObj.childNodes[a].childNodes[7].text : "";
//                var points = new VELatLong(lat,lng);
//                
//                var infobox = "<p><img src='http://whc.unesco.org/uploads/states/small/" + flags[0] + ".gif' style='border: solid 1px #000000;'> ";
//                infobox += (flags[1]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[1] + ".gif' style='border: solid 1px #000000;'> " : "";
//                infobox += (flags[2]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[2] + ".gif' style='border: solid 1px #000000;'> " : "";
//                infobox += (flags[3]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[3] + ".gif' style='border: solid 1px #000000;'> " : "";
//                infobox += (flags[4]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[4] + ".gif' style='border: solid 1px #000000;'> " : "";
//                infobox += (flags[5]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[5] + ".gif' style='border: solid 1px #000000;'> " : "";
//                infobox += (flags[6]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[6] + ".gif' style='border: solid 1px #000000;'> " : "";
//                infobox += (flags[7]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[7] + ".gif' style='border: solid 1px #000000;'> " : "";
//                infobox += (flags[8]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[8] + ".gif' style='border: solid 1px #000000;'> " : "";
//                infobox += (flags[9]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[9] + ".gif' style='border: solid 1px #000000;'> " : "";
//                infobox += (flags[10]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[10] + ".gif' style='border: solid 1px #000000;'> " : "";
//                infobox += (flags[11]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[11] + ".gif' style='border: solid 1px #000000;'> " : "";
//                infobox += "</p><div id='PushpinBalloonContainer' style='margin-left:0;'><table border='0' cellpadding='0' cellspacing='0'><tr><td align='left' valign='top' style='font-family: arial; font-size: 0.7em; color: #000000; padding-right:6px; text-align:justify;'><strong style='font-family:arial;'>Brief Description</strong><br />" + description + "</td><td align='right' valign='top' style='padding-right:2px;'><img src='" + image + "' style='border: solid 1px #979797;' align='right'></td></tr></table></div>";
//                
//                var customIcon = "<div style='padding: 1px 1px 1px 1px; margin: 0px; border: 1px solid #EEE; font-weight:bold; background-color: #b0e3fa; color: #00156e; text-decoration:none; font-family:verdana; font-size:11px;'>" + b + "</div>";
//                  b++;
//                
//                var shape = new VEShape(VEShapeType.Pushpin, points);
//                shape.SetCustomIcon(customIcon);
//                shape.SetTitle(title);
//                shape.SetDescription(infobox);
//                map.ClearInfoBoxStyles();
//                map.AddShape(shape);
////                var pin = map.AddPushpin(points);
////                pin.SetCustomIcon(customIcon);
////                pin.SetTitle(title);
////                pin.SetDescription(infobox);    
//            }
//            setTimeout('document.getElementById("inprogress_img").style.cssText = "display:none;"', 2000);
//            setTimeout('document.getElementById("loader").style.cssText = "display:none;"', 2000);
//        }
//    }
}



