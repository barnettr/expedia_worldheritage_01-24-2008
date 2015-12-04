// JScript File
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
    <title></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <script type="text/javascript" src="http://dev.virtualearth.net/mapcontrol/mapcontrol.ashx?v=6"></script>
    <script type="text/javascript">

        var horizontalOffset = 200;
        var verticalOffset = 100;
        
        function GetMap()
        {
            map = new VEMap('myMap');
            map.LoadMap();
            
            // Important!
            map.ClearInfoBoxStyles();

            map.AttachEvent("onmouseover", OnMouseOver);
        }

        function OnMouseOver(e)
        {
            if (e.elementID)
            {
                var hoveredShape = map.GetShapeByID(e.elementID);

                if (hoveredShape)
                {
                    var ll = hoveredShape.GetIconAnchor();
                    var hookPoint = map.LatLongToPixel(ll);
                    hookPoint.x += horizontalOffset;
                    hookPoint.y += verticalOffset;
                    
                    map.ShowInfoBox(hoveredShape, hookPoint);
                    
                    // VERY IMPORTANT TO RETURN TRUE HERE!!! Otherwise default implementation will show
                    // the infobox on the default location.
                    return true;
                }
            }
        }

        function AddPushpin()
        {
            var latlong = map.GetCenter();
            var shape = new VEShape(VEShapeType.Pushpin, latlong);
            shape.SetTitle('My pushpin');
            // toFixed simply reduces the number of digits for better reading.
            shape.SetDescription('This pushpin is at:<br>Latitude: ' + latlong.Latitude.toFixed(5) + '<br>Longitude: ' + latlong.Longitude.toFixed(5));
            map.AddShape(shape);        
        }

    </script>
</head>
<body onload="GetMap();">
<div id='myMap' style="position:relative; width:800px; height:800px;"></div>
<div><a href='#' onclick='AddPushpin();'>Add Pushpin to the center of the map.</a></div>
</body>
</html>




   <script type="text/javascript" src="http://dev.virtualearth.net/mapcontrol/mapcontrol.ashx?v=6"></script>
    <script type="text/javascript">

        var infobox = null;
        var shape = null;
        
        function GetMap()
        {
            map = new VEMap('myMap');
            map.LoadMap();
            
            var latlong = map.GetCenter();
            shape = new VEShape(VEShapeType.Pushpin, latlong);
            shape.SetTitle('My pushpin');
            shape.SetDescription('This pushpin is at:<br>Latitude: ' + latlong.Latitude.toFixed(5) + '<br>Longitude: ' + latlong.Longitude.toFixed(5));
            map.AddShape(shape);    
        }

        function ShowInfobox()
        {                
            map.ShowInfoBox(shape);

            infobox = GetInfoBoxReference();
            
            if (infobox)
            {
                MoveInfoBox();
            }
        }
        
        function MoveInfoBox()
        {
            // When we are at the top, stop the animation.
            if (parseInt(infobox.style.top) == 0)
                return;
                
            infobox.style.top = (parseInt(infobox.style.top) - 1) + (document.all?"":"px;");
            
            setTimeout("MoveInfoBox()", 50);
        }
        
        // Gets a reference to the infobox, looking for its class name within
        // the child nodes of body. This function is not reliable at all...
        function GetInfoBoxReference()
        {
            for (var index in document.body.childNodes)
            {
                var a = document.body.childNodes[index];

                if ( (a.className == "customInfoBox-with-leftBeak") ||
                    (a.className == "customInfoBox-with-rightBeak") ||
                    (a.className == "customInfoBox-noBeak") ||
                    (a.className == "ero ero-leftBeak") ||
                    (a.className == "ero ero-rightBeak") ||
                    (a.className == "ero ero-noBeak")
                    )
                {
                    // It seems to be what we are looking for... it seems...
                    return a;
                }
            }
            
            return null;
        }

    </script>








var mouseOverPinID = map.GetShapeByID(e.elementID);
                //var mouseOverPinX = document.getElementById(mouseOverPinID).style.posTop;
                //alert(mouseOverPinX);
                //.GetID()
                
                
                var horizontalOffset = 0;
                var verticalOffset = 0;
                //alert(mouseOverPinID);
                if (mouseOverPinID)
                {
                    //alert("inside conditional statement");
                    var latLon = mouseOverPinID.GetIconAnchor();
                    //alert("ll = "+latLon);
                    var hookPoint = map.LatLongToPixel(latLon);
                    hookPoint.x += horizontalOffset;
                    hookPoint.y += verticalOffset;
                    if (hookPoint.y > 180)
                    {
                        hookPoint.y += 260;
                    }
                    //alert(hookPoint.y);
                    
                    map.ShowInfoBox(mouseOverPinID, hookPoint);
                    return true;
   
               }
               
               
               
               
               
               
               
        
        function MoveInfoBox()
        {
            // When we are at the top, stop the animation.
            if (parseInt(infobox.style.posTop) == 0)
                return;
                
            infobox.style.posTop = (parseInt(infobox.style.posTop) - 1) + (document.all?"":"-50px;");
            
            setTimeout("MoveInfoBox()", 0);
        }
        
        // Gets a reference to the infobox, looking for its class name within
        // the child nodes of body. This function is not reliable at all...
        function GetInfoBoxReference()
        {
            for (var index in document.body.childNodes)
            {
                var a = document.body.childNodes[index];

                if ( (a.className == "customInfoBox-with-leftBeak") ||
                    (a.className == "customInfoBox-with-rightBeak") ||
                    (a.className == "customInfoBox-noBeak") ||
                    (a.className == "ero ero-leftBeak") ||
                    (a.className == "ero ero-rightBeak") ||
                    (a.className == "ero ero-noBeak")
                    )
                {
                    // It seems to be what we are looking for... it seems...
                    return a;
                }
            }
            
            return null;
        }
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        latlong.Latitude.toFixed(5) + '<br>Longitude: ' + latlong.Longitude.toFixed(5
        
                 var mouseOverPinID = map.GetShapeByID(e.elementID);
                //var mouseOverPinX = document.getElementById(mouseOverPinID).style.posTop;
                //alert(mouseOverPinX);
                //.GetID()
                
                var horizontalOffset = 0;
                var verticalOffset = 0;
               
                if (mouseOverPinID)
                {
                    //alert("inside conditional statement");
                    var latLon = mouseOverPinID.GetIconAnchor();
                    var hookPoint = map.LatLongToPixel(latLon);
                    hookPoint.x += horizontalOffset;
                    hookPoint.y += verticalOffset;
                    alert(hookPoint.x);
                    alert(hookPoint.y);
                    
                    map.ShowInfoBox(mouseOverPinID, hookPoint);
                    return true;
   
               }