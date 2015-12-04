function loadHandler () {
  xmlProcessor(this);
}	
	
function loadXML()
{ 
//	var xslStylesheet;
//    var myDOM;
//    var xmlDoc;
//    var xsltProcessor = new XSLTProcessor(); 
//    
//    var myXMLHTTPRequest = new XMLHttpRequest(); 
//    myXMLHTTPRequest.open("GET","worldheritage_grid.xslt", false);
//    myXMLHTTPRequest.send(null);
//    xslStylesheet = myXMLHTTPRequest.responseXML;
//    xsltProcessor.importStylesheet(xslStylesheet); 
//    
//    myXMLHTTPRequest = new XMLHttpRequest();
//    myXMLHTTPRequest.open("GET", sFile, false);
//    myXMLHTTPRequest.send(null);
//    xmlDoc = myXMLHTTPRequest.responseXML; 
//    xsltProcessor.setParameter("worldheritage","Title",""); 
//    var ownerDocument=document.implementation.createDocument("", "test", null); 
//    var newFragment=xsltProcessor.transformToFragment(xmlDoc,document);
//    document.body.appendChild( newFragment); 
	
	//var processor = new XSLTProcessor();
	var sFile = document.getElementById("sFile").value;
	
	if (document.implementation && document.implementation.createDocument)
	{
		xmlDoc = document.implementation.createDocument("","",null);
		xmlDoc.load(sFile);
		xmlDoc.addEventListener('load', loadHandler, false);
	}
}

function xmlProcessor(xmlDoc) 
{
    if(!xmlDoc) 
    {
        alert("The XML document did not load!");
        return false;
    }
    
	var selStyle = "r";
	var selMode = VEMapMode.Mode2D;
	var fixed = 0;
	var showSwitch = 1;
	var zoom = 3;
	var map = null;
	
	var Pushpin_Region = "";
	var Pushpin_State = "";
	
	var Pushpin_Region = document.getElementById("sRegion").value;
    var Pushpin_State = document.getElementById("sState").value;
	
	currentPage = document.getElementById("paramCurrentPage").value;
	pageSize = document.getElementById("paramPageSize").value;
	sortCol = document.getElementById("paramSortCol").value;
	sortDir = document.getElementById("paramSortDir").value;
	sortType = document.getElementById("paramSortType").value;
	
	if (Pushpin_Region != "" && Pushpin_State == "")
	{
		switch (Pushpin_Region)
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
        
        function OnMouseOver(e)
        {
            if (e.elementID)
            {
                y = event.pageY;
                x = event.pageX;
                var pixel = new VEPixel((x-380), (y+150));
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
        
        map = new VEMap('myMap');
        map.LoadMap(latLon, zoom, selStyle, fixed, selMode, showSwitch);
        map.SetScaleBarDistanceUnit(VEDistanceUnit.Miles);
        map.ClearInfoBoxStyles();
        
		
		var rowNodes = xmlDoc.getElementsByTagName('row');
	    var b = 1;
		for (var a = 0; a < rowNodes.length; a++)
	    {
			if (rowNodes[a].tagName == "row" && rowNodes[a].childNodes[7].text == Pushpin_Region)
	        {
				if ((a >= 0 + (currentPage -1) * pageSize) && (a < (0 + currentPage * pageSize)))
				{
				    var flags = rowNodes[a].childNodes[5].text.split(",");
				    var location = rowNodes[a].childNodes[9].text;
				    var image = "http://whc.unesco.org/uploads/sites/site_" + rowNodes[a].getAttribute("id") + ".jpg";
	                var lat = parseFloat(rowNodes[a].getAttribute("lat"));
	                var lng = parseFloat(rowNodes[a].getAttribute("long"));
	                var id = rowNodes[a].getAttribute("id");
	                var suitcase = "<img src='suitcase.gif' style='border:none;'";
	                var description = rowNodes[a].childNodes[15].text;
	                points = new VELatLong(lat,lng);
    	            	            
	                var infobox = "<p>";
                    for (x = 0; x < flags.length; x++ )
                    {
                        infobox += (flags[x]) ? "<img src='http://whc.unesco.org/uploads/states/small/" + flags[x] + ".gif' style='border: solid 1px #000000;'> " : "";
                    }
	                infobox += "</span><div id='PushpinBalloonContainer' style='margin-top:0;'><table border='0' cellpadding='0' cellspacing='0'><tr><td align='left' valign='top' style='font-family: arial; font-size: 1.0em; color: #000000; padding-right:8px; text-align:justify;'><strong style='font-family:arial;'>Brief Description</strong><br />" + description + "</td><td align='right' valign='top' style='padding-right:2px;'><img src='" + image + "' style='border: solid 1px #979797;' align='right'></td></tr></table></div>";
	                
	                if (currentPage == 1)
                    {
                        var title = "<table border='0' cellpadding='0' cellspacing='0' style='width:100%'><tr><td align='left' valign='top' style='width:16px;'><div class='counter'>" + b + "</div></td><td align='left' valign='top' style='text-align:left; padding-left:4px;'><a href='' style='text-decoration:none; border:0px; font-weight:bold; text-align:left;'>" + rowNodes[a].childNodes[13].text + "</a></td><td align='right' valign='top' style='width:41px; height:34px; padding-left:4px;'>" + suitcase + "</td></tr></table>";
                        var customIcon = "<div id='pin_" + id + "' style='padding: 1px 1px 1px 1px; margin: 0px; border: 1px solid #EEE; font-weight:bold; background-color: #b0e3fa; color: #00156e; text-decoration:none; font-family:verdana; font-size:11px;'>" + b + "</div>";
                    }
                    else
                    {
                        var title = "<table border='0' cellpadding='0' cellspacing='0' style='width:100%'><tr><td align='left' valign='top' style='width:16px;'><div class='counter'>" + (b+((currentPage-1) * pageSize)) + "</div></td><td align='left' valign='top' style='text-align:left; padding-left:4px;'><a href='' style='text-decoration:none; border:0px; font-weight:bold; text-align:left;'>" + rowNodes[a].childNodes[13].text + "</a></td><td align='right' valign='top' style='width:41px; height:34px; padding-left:4px;'>" + suitcase + "</td></tr></table>";
                        var customIcon = "<div id='pin_"+id+"' style='padding: 1px 1px 1px 1px; margin: 0px; border: 1px solid #EEE; font-weight:bold; background-color: #b0e3fa; color: #00156e; text-decoration:none; font-family:verdana; font-size:11px;'>" + (b+((currentPage-1) * pageSize)) + "</div>";
                    }
	                  b++
    	          
	                var shape = new VEShape(VEShapeType.Pushpin, points);
                    shape.SetCustomIcon(customIcon);
                    shape.SetTitle(title);
                    map.AttachEvent("onmouseover", OnMouseOver);
                    map.AttachEvent("onmouseout", OnMouseOut);
                    shape.SetDescription(infobox);
                    map.AddShape(shape);
                }
	        }
	    }
	    setTimeout('document.getElementById("inprogress_img").style.cssText = "display:none;"', 90);
        setTimeout('document.getElementById("loader").style.cssText = "display:none;"', 90);
	}
	else if (Pushpin_Region != "" && Pushpin_State != "")  // Two dropdown boxes: Region & State
	{
		var rowNodes = xmlDoc.getElementsByTagName('row');
	    var b = 1;
		
        for (var x = 0; x < rowNodes.length; x++)
        {
            if (rowNodes[x].tagName == "row" && rowNodes[x].childNodes[7].text == Pushpin_Region && rowNodes[x].childNodes[11].text == Pushpin_State)
            {
                var lat = parseFloat(rowNodes[x].getAttribute("lat"));
                var lng = parseFloat(rowNodes[x].getAttribute("long"));
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
		
		for (var a = 0; a < rowNodes.length; a++)
	    {
			if (rowNodes[a].tagName == "row" && rowNodes[a].childNodes[7].text == Pushpin_Region && rowNodes[a].childNodes[11].text == Pushpin_State)
	        {
				var flags = rowNodes[a].childNodes[5].text.split(",");
				var location = rowNodes[a].childNodes[9].text;
				var image = "http://whc.unesco.org/uploads/sites/site_" + rowNodes[a].getAttribute("id") + ".jpg";
	            var lat = parseFloat(rowNodes[a].getAttribute("lat"));
	            var lng = parseFloat(rowNodes[a].getAttribute("long"));
	            var suitcase = "<img src='suitcase.gif' style='border:none;'";
                var title = "<table border='0' cellpadding='0' cellspacing='0' style='width:100%'><tr><td align='left' valign='top' style='width:16px;'><div class='counter'>" + b + "</div></td><td align='left' valign='top' style='text-align:left; padding-left:4px;'><a href='' style='text-decoration:none; border:0px; font-weight:bold; text-align:left;'>" + rowNodes[a].childNodes[13].text + "</a></td><td align='right' valign='top' style='width:41px; height:34px; padding-left:4px;'>" + suitcase + "</td></tr></table>";
	            var description = rowNodes[a].childNodes[15].text;
	            points = new VELatLong(lat,lng);
	            	            
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
	            infobox += "</p><div id='PushpinBalloonContainer'><table border='0' cellpadding='0' cellspacing='0'><tr><td align='left' valign='top' style='font-family: arial; font-size: 1.1em; color: #000000; padding-right:6px; text-align:justify;'><strong style='font-family:arial;'>Brief Description</strong><br />" + description + "</td><td align='right' valign='top' style='padding-right:2px;'><img src='" + image + "' style='border: solid 1px #979797;' align='right'></td></tr></table></div>";
	            var customIcon = "<div style='padding: 1px 1px 1px 1px; margin: 0px; border: 1px solid #EEE; font-weight:bold; background-color: #b0e3fa; color: #00156e; text-decoration:none; font-family:verdana; font-size:11px;'>" + b + "</div>";
	              b++
	          
	            var shape = new VEShape(VEShapeType.Pushpin, points);
                shape.SetCustomIcon(customIcon);
                shape.SetTitle(title);
                shape.SetDescription(infobox);
                map.AddShape(shape);
//	            var pin = map.AddPushpin(points);
//	            pin.SetCustomIcon(customIcon);
//	            pin.SetTitle(title);
//	            pin.SetDescription(infobox);
	        }
	    }
	    setTimeout('document.getElementById("inprogress_img").style.cssText = "display:none;"', 90);
        setTimeout('document.getElementById("loader").style.cssText = "display:none;"', 90);
	}
//	else if (Pushpin_Group != "" && Pushpin_Region != "" && Pushpin_Location != "") // All three dropdown boxes selected by user: Region & State & Location
//	{
//	    var rowNodes = xmlDoc.getElementsByTagName('row');
//	    var b = 1;
//	    
//        for (var x = 0; x < rowNodes.length; x++)
//        {
//            if (rowNodes[x].tagName == "row" && rowNodes[x].childNodes[7].text == Pushpin_Group && rowNodes[x].childNodes[9].text == Pushpin_Location && rowNodes[x].childNodes[11].text == Pushpin_Region )
//            {
//                var lat = parseFloat(rowNodes[x].getAttribute("lat"));
//                var lng = parseFloat(rowNodes[x].getAttribute("long"));
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
//		for (var a = 0; a < rowNodes.length; a++)
//	    {
//			if (rowNodes[a].tagName == "row" && rowNodes[a].childNodes[7].text == Pushpin_Group && rowNodes[a].childNodes[9].text == Pushpin_Location && rowNodes[a].childNodes[11].text == Pushpin_Region)
//	        {
//				var flags = rowNodes[a].childNodes[5].text.split(",");
//				var location = rowNodes[a].childNodes[9].text;
//				var image = "http://whc.unesco.org/uploads/sites/site_" + rowNodes[a].getAttribute("id") + ".jpg";
//	            var lat = parseFloat(rowNodes[a].getAttribute("lat"));
//	            var lng = parseFloat(rowNodes[a].getAttribute("long"));
//	            var suitcase = "<img src='suitcase.gif' style='border:none;'";
//                var title = "<table border='0' cellpadding='0' cellspacing='0' style='width:100%'><tr><td align='left' valign='top' style='width:16px;'><div class='counter'>" + b + "</div></td><td align='left' valign='top' style='text-align:left; padding-left:4px;'><a href='' style='text-decoration:none; border:0px; font-weight:bold; text-align:left;'>" + rowNodes[a].childNodes[13].text + "</a></td><td align='right' valign='top' style='width:41px; height:34px; padding-left:4px;'>" + suitcase + "</td></tr></table>";
//	            var description = rowNodes[a].childNodes[15].text;
//	            points = new VELatLong(lat,lng);
//	            	            
//	            var infobox = "<p><img src='http://whc.unesco.org/uploads/states/small/" + flags[0] + ".gif' style='border: solid 1px #000000;'> ";
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
//	            infobox += "</p><div id='PushpinBalloonContainer'><table border='0' cellpadding='0' cellspacing='0'><tr><td align='left' valign='top' style='font-family: arial; font-size: 1.1em; color: #000000; padding-right:6px; text-align:justify;'><strong style='font-family:arial;'>Brief Description</strong><br />" + description + "</td><td align='right' valign='top' style='padding-right:2px;'><img src='" + image + "' style='border: solid 1px #979797;' align='right'></td></tr></table></div>";
//	            var customIcon = "<div style='padding: 1px 1px 1px 1px; margin: 0px; border: 1px solid #EEE; font-weight:bold; background-color: #b0e3fa; color: #00156e; text-decoration:none; font-family:verdana; font-size:11px;'>" + b + "</div>";
//	              b++
//	          
//	            var shape = new VEShape(VEShapeType.Pushpin, points);
//                shape.SetCustomIcon(customIcon);
//                shape.SetTitle(title);
//                shape.SetDescription(infobox);
//                map.AddShape(shape);
////	            var pin = map.AddPushpin(points);
////	            pin.SetCustomIcon(customIcon);
////	            pin.SetTitle(title);
////	            pin.SetDescription(infobox);
//	        }
//	    }
//	    setTimeout('document.getElementById("inprogress_img").style.cssText = "display:none;"', 90);
//        setTimeout('document.getElementById("loader").style.cssText = "display:none;"', 90);
//	}
}
