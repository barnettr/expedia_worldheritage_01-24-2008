<% 
    If InStr(1, Request.ServerVariables("HTTP_USER_AGENT"), "MSIE") Then

%>
	<link title="combined" rel="stylesheet" type="text/css" media="screen" href="indexview.css" />
<% 
    Else
%>
    <link title="combined" rel="stylesheet" type="text/css" media="screen" href="indexview_ff.css" />
<% 
    End If
%>

	<style type="text/css">
	    .linkHolder {font-family: arial, helvetica, sans serif;font-size: 11px;font-weight:normal;cursor:pointer;}
	    a.linkHolder{color:#4985c7 !important;text-decoration:none;}
        a.linkHolder:visited{color:#4985c7 !important;text-decoration:none;}
        a.linkHolder:hover{color:#cc6600 !important;text-decoration:underline;} 
        a.linkHolderHover{color:#cc6600 !important;text-decoration:underline;}     
	</style>
	<script type="text/javascript" src="http://dev.virtualearth.net/mapcontrol/mapcontrol.ashx?v=6"></script>
	<script language="javascript" type="text/javascript">
	
	
		var map = null;
		var latLon = new VELatLong(34.031199, 118.144196);
		zoom = 1;
		var pinid = 0;
		var selStyle = "r";  // a, h, r
		var selMode = VEMapMode.Mode2D;
		var fixed = 0;
        var showSwitch = 1;
	    
	    var arrPin = [];
        
        var pin0 = new Object();
        pin0.locid = "326";
        pin0.region = "Arab States";
        pin0.site = "Petra";
        pin0.location = "Ma’an Governorate";
        pin0.state = "Jordan";
        pin0.desc = "Inhabited since prehistoric times, this Nabataean caravan-city, situated between the Red Sea and the Dead Sea, was an important crossroads between Arabia, Egypt and Syria-Phoenicia. Petra is half-built, half-carved into the rock, and is surrounded by mountains riddled with passages and gorges. It is one of the world's most famous archaeological sites, where ancient Eastern traditions blend with Hellenistic architecture.";
        pin0.img = "http://whc.unesco.org/uploads/sites/site_326.jpg";
        pin0.lat = "30.33056000";
        pin0.lng = "35.44333000";
        pin0.customIcon = "<div class='pinImgOff' onmouseover='this.className=\"pinImgOn\";' onmouseout='this.className=\"pinImgOff\";'>1</div>";
        arrPin[pinid] = pin0;
        pinid++;
        
        var pin1 = new Object();
        pin1.locid = "274";326274
        pin1.region = "Latin America";
        pin1.site = "Historic Sanctuary of Machu Picchu";
        pin1.location = "Left bank of the Vilcanota"; //river, between its tributaries the Kusichaca and Aobamba rivers, in the department of Cuzco, Urubamba Province, District of Macchu Picchu";
        pin1.state = "Peru";
        pin1.desc = "Machu Picchu stands 2,430 m above sea-level, in the middle of a tropical mountain forest, in an extraordinarily beautiful setting. It was probably the most amazing urban creation of the Inca Empire at its height; its giant walls, terraces and ramps seem as if they have been cut naturally in the continuous rock escarpments. The natural setting, on the eastern slopes of the Andes, encompasses the upper Amazon basin with its rich diversity of flora and fauna.";
        pin1.img = "http://whc.unesco.org/uploads/sites/site_274.jpg";
        pin1.lat = "-13.11666667";
        pin1.lng = "-72.58333333";
        pin1.customIcon = "<div class='pinImgOff' onmouseover='this.className=\"pinImgOn\";' onmouseout='this.className=\"pinImgOff\";'>2</div>";
        arrPin[pinid] = pin1;
        pinid++;
        
        var pin2 = new Object();
        pin2.locid = "307";
        pin2.region = "North America";
        pin2.site = "Statue of Liberty";
        pin2.location = "New York, New York";
        pin2.state = "United States of America";
        pin2.desc = "Made in Paris by the French sculptor Bartholdi, in collaboration with Gustave Eiffel (who was responsible for the steel framework), this towering monument to liberty was a gift from France on the centenary of American independence in 1886. Standing at the entrance to New York Harbour, it has welcomed millions of immigrants to the United States ever since.";
        pin2.img = "http://whc.unesco.org/uploads/sites/site_307.jpg";
        pin2.lat = "40.68944444";
        pin2.lng = "-74.04472222";
        pin2.customIcon = "<div class='pinImgOff' onmouseover='this.className=\"pinImgOn\";' onmouseout='this.className=\"pinImgOff\";'>3</div>";
        arrPin[pinid] = pin2;
        pinid++;
        
        var pin3 = new Object();
        pin3.locid = "616";
        pin3.region = "Europe";
        pin3.site = "Historic Centre of Prague";
        pin3.location = "Prague (Hlavní mesto Praha)";
        pin3.state = "Czech Republic";
        pin3.desc = "Built between the 11th and 18th centuries, the Old Town, the Lesser Town and the New Town speak of the great architectural and cultural influence enjoyed by this city since the Middle Ages. The many magnificent monuments, such as Hradcani Castle, St Vitus Cathedral, Charles Bridge and numerous churches and palaces, built mostly in the 14th century under the Holy Roman Emperor, Charles IV.";
        pin3.img = "http://whc.unesco.org/uploads/sites/site_616.jpg";
        pin3.lat = "50.08972000";
        pin3.lng = "14.41944000";
        pin3.customIcon = "<div class='mapArrow' onmouseover='this.className=\"mapArrowOn\";' onmouseout='this.className=\"mapArrow\";'></div>";
        arrPin[pinid] = pin3;
        pinid++;
        
        var pin4 = new Object();
        pin4.locid = "980";
        pin4.region = "Europe";
        pin4.site = "Historic and Architectural Complex of the Kazan Kremlin";
        pin4.location = "City of Kazan, Republic of Tatarstan";
        pin4.state = "Russian Federation";
        pin4.desc = "Built on an ancient site, the Kazan Kremlin dates from the Muslim period of the Golden Horde and the Kazan Khanate. It was conquered by Ivan the Terrible in 1552 and became the Christian See of the Volga Land. The only surviving Tatar fortress in Russia and an important place of pilgrimage, the Kazan Kremlin consists of an outstanding group of historic buildings dating from the 16th to 19th centuries, integrating remains of earlier structures of the 10th to 16th centuries.";
        pin4.img = "http://whc.unesco.org/uploads/sites/site_980.jpg";
        pin4.lat = "55.79111111";
        pin4.lng = "49.09500000";
        pin4.customIcon = "<div class='mapArrow' onmouseover='this.className=\"mapArrowOn\";' onmouseout='this.className=\"mapArrow\";'></div>";
        arrPin[pinid] = pin4;
        pinid++;
        
        var pin5 = new Object();
        pin5.locid = "1021";
        pin5.region = "Africa";
        pin5.site = "Tsodilo";
        pin5.location = "The Ngamiland District, north-west Botswana";
        pin5.state = "Botswana";
        pin5.desc = "With one of the highest concentrations of rock art in the world, Tsodilo has been called the 'Louvre of the Desert'. Over 4,500 paintings are preserved in an area of only 10 km2 of the Kalahari Desert. The archaeological record of the area gives a chronological account of human activities and environmental changes over at least 100,000 years. Local communities in this hostile environment respect Tsodilo as a place of worship frequented by ancestral spirits.";
        pin5.img = "http://whc.unesco.org/uploads/sites/site_1021.jpg";
        pin5.lat = "-18.75000000";
        pin5.lng = "21.73333333";
        pin5.customIcon = "<div class='mapArrow' onmouseover='this.className=\"mapArrowOn\";' onmouseout='this.className=\"mapArrow\";'></div>";
        arrPin[pinid] = pin5;
        pinid++;
        
        var pin6 = new Object();
        pin6.locid = "881";
        pin6.region = "Asia";
        pin6.site = "Temple of Heaven: an Imperial Sacrificial Altar in Beijing";
        pin6.location = "Tiantan Park, Beijing";
        pin6.state = "China";
        pin6.desc = "The Temple of Heaven, founded in the first half of the 15th century, is a dignified complex of fine cult buildings set in gardens and surrounded by historic pine woods. In its overall layout and that of its individual buildings, it symbolizes the relationship between earth and heaven – the human world and God's world – which stands at the heart of Chinese cosmogony, and also the special role played by the emperors within that relationship.";
        pin6.img = "http://whc.unesco.org/uploads/sites/site_881.jpg";
        pin6.lat = "39.84555556";
        pin6.lng = "116.44472220";
        pin6.customIcon = "<div class='mapArrow' onmouseover='this.className=\"mapArrowOn\";' onmouseout='this.className=\"mapArrow\";'></div>";
        arrPin[pinid] = pin6;
        pinid++;
        
        var pin7 = new Object();
        pin7.locid = "814";
        pin7.region = "Caribbean";
        pin7.site = "Morne Trois Pitons National Park";
        pin7.location = "South central part of the island";
        pin7.state = "Dominica";
        pin7.desc = "Luxuriant natural tropical forest blends with scenic volcanic features of great scientific interest in this national park centred on the 1,342-m-high volcano known as Morne Trois Pitons. With its precipitous slopes and deeply incised valleys, 50 fumaroles, hot springs, three freshwater lakes, a &amp;#39;boiling lake&amp;#39; and five volcanoes, located on the park&amp;#39;s nearly 7,000 ha, together with the richest biodiversity in the Lesser Antilles, Morne Trois Pitons National Park presents a rare combination of natural features of World Heritage value.";
        pin7.img = "http://whc.unesco.org/uploads/sites/site_814.jpg";
        pin7.lat = "15.26666667";
        pin7.lng = "-61.28333333";
        pin7.customIcon = "<div class='mapArrow' onmouseover='this.className=\"mapArrowOn\";' onmouseout='this.className=\"mapArrow\";'></div>";
        arrPin[pinid] = pin7;
        pinid++;
        
        var pin8 = new Object();
        pin8.locid = "590";
        pin8.region = "Asia";
        pin8.site = "Dong Phayayen-Khao Yai Forest Complex";
        pin8.location = "Provinces of Saraburi";//, Nakhon Nayok, Nakhon Rachisima, Prachinburi, Srakaew and Burirum";
        pin8.state = "Thailand";
        pin8.desc = "The Dong Phayayen-Khao Yai Forest Complex spans 230 km between Ta Phraya National Park on the Cambodian border in the east, and Khao Yai National Park in the west. The site is home to more than 800 species of fauna, including 112 mammal species (among them two species of gibbon), 392 bird species and 200 reptile and amphibian species. It is internationally important for the conservation of globally threatened and endangered mammal, bird and reptile species, among them 19 that are vulnerable, four that are endangered, and one that is critically endangered. The area contains substantial and important tropical forest ecosystems, which can provide a viable habitat for the long-term survival of these species.";
        pin8.img = "http://whc.unesco.org/uploads/sites/site_590.jpg";
        pin8.lat = "14.33000000";
        pin8.lng = "102.05000000";
        pin8.customIcon = "<div class='mapArrow' onmouseover='this.className=\"mapArrowOn\";' onmouseout='this.className=\"mapArrow\";'></div>";
        arrPin[pinid] = pin8;
        pinid++;
        
        var pin9 = new Object();
        pin9.locid = "166";
        pin9.region = "Pacific";
        pin9.site = "Sydney Opera House";
        pin9.location = "New South Wales";
        pin9.state = "Australia";
        pin9.desc = "Inaugurated in 1973, the Sydney Opera House is listed as a great architectural work of the 20th century that brings together multiple strands of creativity and innovation, both in architectural form and structural design. A great urban sculpture set in a remarkable waterscape, at the tip a peninsula projecting into Sydney Harbour, the building has had an enduring influence on architecture. The Opera House comprises three groups of interlocking vaulted ‘shells&amp;#39; which roof two main performances halls and a restaurant. These shell-structure are set upon a vast platform and are surrounded by terrace areas that function as pedestrian concourses. In 1957, when the project of the Sydney opera was attributed by an international jury to the then almost unknown Danish architect Jørn Utzon, it marked a radically new and collaborative approach to construction. In listing the building, the Sydney Opera House is recognized as a great artistic monument accessible to society at large.";
        pin9.img = "http://whc.unesco.org/uploads/sites/site_1008.jpg";
        pin9.lat = "-33.85666667";
        pin9.lng = "151.21527778";
        pin9.customIcon = "<div class='mapArrow' onmouseover='this.className=\"mapArrowOn\";' onmouseout='this.className=\"mapArrow\";'></div>";
        arrPin[pinid] = pin9;
        pinid++;
        
        var infobox = null;
    	function GetMap()
        {
	        map = new VEMap('myMap');
	        map.SetDashboardSize(VEDashboardSize.Normal);	
	        map.onLoadMap = MapLoaded;
	        map.LoadMap(latLon,zoom,'r');
	        map.AttachEvent("onmouseover", mouseoverPin);
	        map.AttachEvent("onmouseout", mouseoutPin);	
            map.AttachEvent("onmousewheel", onscrollwheel);

	        map.ClearInfoBoxStyles();
	        AddPushpins();
	        buildLinks();
        }
        
        function MapLoaded(){}
        
        function onscrollwheel(e)
        {
             return true;
        }      
        
        function mouseoverPin(e)
        {
	        if (e.elementID != null)
            {                       
 		        var fl = getLinkID(map.GetShapeByID(e.elementID).GetID());
 		        if(fl != '274' && fl != '307' && fl != '326')
 		        {
		            showLink(fl); 
		        }              
	        }
        }

        function mouseoutPin(e)
        {
	        if (e.elementID != null)
            {
 		        var fl = getLinkID(map.GetShapeByID(e.elementID).GetID());
 		        if(fl != '274' && fl != '307' && fl != '326')
 		        {
		            hideLink(fl);
		        }            
	        }
        }
        
        function getLinkID(val)
        {
	        for (x = 0; x < arrPin.length; x++)
	        {
		        if (arrPin[x].pinid == val)
		        {
			        return arrPin[x].locid;
		        }
	        }
        }
        
        function showLink(id)
        {
            var fl = "fl_"+id;
            document.getElementById(fl).className = "linkHolderHover";
        }

        function hideLink(id)
        {
            var fl = "fl_"+id;
            document.getElementById(fl).className = "linkHolder";
        }        
        
        function AddPushpins()
        {
	        for (x = 0; x < arrPin.length; x++)
	        {
		        var pin = BuildPushpin(arrPin[x].locid,arrPin[x].region,arrPin[x].site,arrPin[x].location,arrPin[x].state,arrPin[x].desc,arrPin[x].img,arrPin[x].lat,arrPin[x].lng,arrPin[x].customIcon);
		        map.AddShape(pin);
		        arrPin[x].pinid = pin.GetID();
	        }
        }
        
        function BuildPushpin(locid,region,site,location,state,desc,img,lat,lng,customIcon)
        {
	        var loc = new VELatLong(lat,lng);
	        var title = "<table border='0' cellpadding='0' cellspacing='0' style='width:100%'><tr><td align='left' valign='top' style='width:16px;'><span style='padding: 1px 3px 1px 3px; margin: 2px; border: 1px solid #EEE; font-weight: bold; background-color: #4985c7; color: #FFF;'>" + locid + "</span></td><td align='left' valign='top' style='text-align:left; padding-left:6px;'><a href='' style='text-decoration:none; border:0px; font-weight:bold; text-align:left; color:#4985c7'>" + site + "</a></td><td align='right' valign='top' style='width:41px; height:34px; padding-left:4px;'></td></tr></table>"
	        
	        var description = "<div style='margin-left:0;'><table border='0' cellpadding='0' cellspacing='0'><tr><td align='right' valign='top' style='padding-right:0px; padding-top:10px;'><img src='" + img + "' style='border: solid 1px #c0c0c1; padding:1px 1px 1px 1px; background-color:#ffffff;' align='left' width='88'></td><td align='left' valign='top' style='font-family:arial; font-size:11; color:#333333; padding-left:6px; padding-top:7px; text-align:left;'><a href='javascript:this.map.HideInfoBox();'><img src='http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_close.gif' align='right' style='margin-left:3px;' id='close'></a><a href='sites.asp?region="+region+"&locid="+locid+"' class='infoBoxLinkOff' onmouseover='this.className=\"infoBoxLinkOn\";' onmouseout='this.className=\"infoBoxLinkOff\";'><b>" + site + "</b></a><div style='margin-top:3px;margin-bottom:3px;'>"+ location +"<br /><div style='margin-top:2px;'>"+ state +"</div></div>" + truncate(desc,region,locid) + "<br /><br /><a href='sites.asp?region="+region+"&locid="+locid+"'><img src='http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_landing_arrow.gif' style='border:none;padding-right:2px;'></a> <a href='sites.asp?region="+region+"&locid="+locid+"' class='sitepageLinkOff' onmouseover='this.className=\"sitepageLinkOn\";' onmouseout='this.className=\"sitepageLinkOff\";' style='vertical-align:top;'>Explore this site</a></td></tr></table></div>";
	        if (navigator.userAgent.indexOf("Firefox") != -1)
	        {
	            var description = "<div style='margin-left:0;'><table border='0' cellpadding='0' cellspacing='0'><tr><td align='right' valign='top' style='padding-top:10px;'><img src='" + img + "' style='border: solid 1px #c0c0c1; padding:1px 1px 1px 1px; background-color:#ffffff;' align='left' width='88'></td><td align='left' valign='top' style='font-family:arial; font-size:11px; color:#333333; padding-left:6px; padding-top:7px; padding-right:8px; text-align:left;'><a href='javascript:this.map.HideInfoBox();'><img src='http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_close.gif' align='right' style='margin-left:3px;padding-right:6px;'></a><a href='sites.asp?region="+region+"&locid="+locid+"' class='infoBoxLinkOff' onmouseover='this.className=\"infoBoxLinkOn\";' onmouseout='this.className=\"infoBoxLinkOff\";'><b>" + site + "</b></a><div style='margin-top:3px;margin-bottom:3px;'>"+ location +"<br /><div style='margin-top:2px;'>"+ state +"</div></div>" + truncate(desc,region,locid) + "<br /><br /><a href='sites.asp?region="+region+"&locid="+locid+"'><img src='http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_landing_arrow.gif' style='border:none;padding-right:2px;'></a> <a href='sites.asp?region="+region+"&locid="+locid+"' class='sitepageLinkOff' onmouseover='this.className=\"sitepageLinkOn\";' onmouseout='this.className=\"sitepageLinkOff\";' style='vertical-align:top;'>Explore this site</a></td></tr></table></div>";
	        }
	        if (navigator.userAgent.indexOf("MSIE 7") != -1)
            {
                var description = "<div style='margin-left:0;'><table border='0' cellpadding='0' cellspacing='0'><tr><td align='right' valign='top' style='padding-right:0px; padding-top:10px;'><img src='" + img + "' style='border: solid 1px #c0c0c1; padding:1px 1px 1px 1px; background-color:#ffffff;' align='left' width='88'></td><td align='left' valign='top' style='font-family:arial; font-size:11px; color:#333333; padding-left:6px; padding-top:7px; text-align:left;'><a href='javascript:this.map.HideInfoBox();'><img src='http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_close.gif' align='right' style='margin-left:3px;padding-right:6px;'></a><a href='sites.asp?region="+region+"&locid="+locid+"' class='infoBoxLinkOff' onmouseover='this.className=\"infoBoxLinkOn\";' onmouseout='this.className=\"infoBoxLinkOff\";'><b>" + site + "</b></a><div style='margin-top:3px;margin-bottom:3px;'>"+ location +"<br /><div style='margin-top:2px;'>"+ state +"</div></div>" + truncate(desc,region,locid) + "<br /><br /><a href='sites.asp?region="+region+"&locid="+locid+"'><img src='http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_landing_arrow.gif' style='border:none;padding-right:2px;'></a> <a href='sites.asp?region="+region+"&locid="+locid+"' class='sitepageLinkOff' onmouseover='this.className=\"sitepageLinkOn\";' onmouseout='this.className=\"sitepageLinkOff\";' style='vertical-align:top;'>Explore this site</a></td></tr></table></div>";
            }
	        var homeIcon = customIcon;
        	
	        var shape = new VEShape(VEShapeType.Pushpin, loc);
	        shape.SetCustomIcon(homeIcon);
	        shape.SetTitle(title);
	        shape.SetDescription(description);
	        return shape;
        }
        function buildLinks(locid,region,site,location,state,desc)
        {
            var holdingCell = document.getElementById("linkHolder");
            
            var spacer1 = document.createElement("span");
            spacer1.style.cssText = "color:#5b5a5a;";
            spacer_tn = document.createTextNode(" | ");
            spacer1.appendChild(spacer_tn);
            var spacer2 = document.createElement("span");
            spacer2.style.cssText = "color:#5b5a5a;";
            spacer_tn = document.createTextNode(" | ");
            spacer2.appendChild(spacer_tn);
            var spacer3 = document.createElement("span");
            spacer3.style.cssText = "color:#5b5a5a;";
            spacer_tn = document.createTextNode(" | ");
            spacer3.appendChild(spacer_tn);
            var spacer4 = document.createElement("span");
            spacer4.style.cssText = "color:#5b5a5a;";
            spacer_tn = document.createTextNode(" | ");
            spacer4.appendChild(spacer_tn);
            var spacer5 = document.createElement("span");
            spacer5.style.cssText = "color:#5b5a5a;";
            spacer_tn = document.createTextNode(" | ");
            spacer5.appendChild(spacer_tn);
            var spacer6 = document.createElement("span");
            spacer6.style.cssText = "color:#5b5a5a;";
            spacer_tn = document.createTextNode(" | ");
            spacer6.appendChild(spacer_tn);
            
            var featureLink1 = document.createElement("a");
            featureLink1.id = "fl_"+pin3.locid;
            featureLink1.className="linkHolder";
            featureLink1.onmouseover = function(){GridShowInfoBox('msftve_1000_200003_10003')};
            featureLink1.onmouseout = function(){GridHideInfoBox('msftve_1000_200003_10003')};
            featureLink1.href = "sites.asp?region=" + pin3.region + "&locid=" + pin3.locid;
            var feature1_tn = document.createTextNode("Prague, " + pin3.state);
            featureLink1.appendChild(feature1_tn);
            holdingCell.appendChild(featureLink1);
            holdingCell.appendChild(spacer1);
            
            var featureLink2 = document.createElement("a");
            featureLink2.id = "fl_"+pin4.locid;
            featureLink2.className="linkHolder";
            featureLink2.onmouseover = function(){GridShowInfoBox('msftve_1000_200004_10004')};
            featureLink2.onmouseout = function(){GridHideInfoBox('msftve_1000_200004_10004')};            
            featureLink2.href = "sites.asp?region=" + pin4.region + "&locid=" + pin4.locid;
            var feature2_tn = document.createTextNode("Kazan Kremlin, Russian Federation");
            featureLink2.appendChild(feature2_tn);
            holdingCell.appendChild(featureLink2);
            holdingCell.appendChild(spacer2);
            
            var featureLink3 = document.createElement("a");
            featureLink3.id = "fl_"+pin5.locid;
            featureLink3.className="linkHolder";
            featureLink3.onmouseover = function(){GridShowInfoBox('msftve_1000_200005_10005')};
            featureLink3.onmouseout = function(){GridHideInfoBox('msftve_1000_200005_10005')};            
            featureLink3.href = "sites.asp?region=" + pin5.region + "&locid=" + pin5.locid;
            var feature3_tn = document.createTextNode("Tsodilo, Botswana");
            featureLink3.appendChild(feature3_tn);
            holdingCell.appendChild(featureLink3);
            holdingCell.appendChild(spacer3);
            
            var featureLink4 = document.createElement("a");
            featureLink4.id = "fl_"+pin6.locid;
            featureLink4.className="linkHolder";
            featureLink4.onmouseover = function(){GridShowInfoBox('msftve_1000_200006_10006')};
            featureLink4.onmouseout = function(){GridHideInfoBox('msftve_1000_200006_10006')};             
            featureLink4.href = "sites.asp?region=" + pin6.region + "&locid=" + pin6.locid;
            var feature4_tn = document.createTextNode("Temple of Heaven, China");
            featureLink4.appendChild(feature4_tn);
            holdingCell.appendChild(featureLink4);
            holdingCell.appendChild(spacer4);
            
            var featureLink5 = document.createElement("a");
            featureLink5.id = "fl_"+pin7.locid;
            featureLink5.className="linkHolder";
            featureLink5.onmouseover = function(){GridShowInfoBox('msftve_1000_200007_10007')};
            featureLink5.onmouseout = function(){GridHideInfoBox('msftve_1000_200007_10007')};             
            featureLink5.href = "sites.asp?region=" + pin7.region + "&locid=" + pin7.locid;
            var feature5_tn = document.createTextNode("Morne Trois Pitons, Dominica");
            featureLink5.appendChild(feature5_tn);
            holdingCell.appendChild(featureLink5);
            holdingCell.appendChild(spacer5);
            
            var featureLink6 = document.createElement("a");
            featureLink6.id = "fl_"+pin8.locid;
            featureLink6.className="linkHolder";
            featureLink6.onmouseover = function(){GridShowInfoBox('msftve_1000_200008_10008')};
            featureLink6.onmouseout = function(){GridHideInfoBox('msftve_1000_200008_10008')};               
            featureLink6.href = "sites.asp?region=" + pin8.region + "&locid=" + pin8.locid;
            var feature6_tn = document.createTextNode("Forest Complex, Thailand");
            featureLink6.appendChild(feature6_tn);
            holdingCell.appendChild(featureLink6);
            holdingCell.appendChild(spacer6);
            
            var featureLink7 = document.createElement("a");
            featureLink7.id = "fl_"+pin9.locid;
            featureLink7.className="linkHolder";
            featureLink7.onmouseover = function(){GridShowInfoBox('msftve_1000_200009_10009')};
            featureLink7.onmouseout = function(){GridHideInfoBox('msftve_1000_200009_10009')};             
            featureLink7.href = "sites.asp?region=" + pin9.region + "&locid=" + pin9.locid;
            var feature7_tn = document.createTextNode("Sydney Opera House, Australia");
            featureLink7.appendChild(feature7_tn);
            holdingCell.appendChild(featureLink7);
        }
        
        function GridShowInfoBox(obj)
        {
	        var pin = map.GetShapeByID(obj);
	        pin.SetZIndex(2000);
	        //var latLon = new VELatLong(pin.GetIconAnchor());
	        //pin.SetIconAnchor(point);
	        if (pin)
	        {
	            var customIcon = "<div class='mapArrowOn' onmouseover='this.className=\"mapArrowOn\";' onmouseout='this.className=\"mapArrow\";'></div>";
	            pin.SetCustomIcon(customIcon);
			    //map.ShowInfoBox(pin);
	        }
        }
        
        function GridHideInfoBox(obj)
        {
	        var pin = map.GetShapeByID(obj);
	        pin.SetZIndex(1000);
	        if (pin)
	        {
	            var customIcon = "<div class='mapArrow' onmouseover='this.className=\"mapArrowOn\";' onmouseout='this.className=\"mapArrow\";'></div>";
	            pin.SetCustomIcon(customIcon);
	            //map.HideInfoBox(pin);
	        }
        }        
        
        function truncate(txt,region,locid)
        {
	        var len = 150;
	        var p = txt;
	        if (p)
	        {
		        var trunc = p;
		        if (trunc.length > len)
		        {
			        trunc = trunc.substring(0, len);
			        trunc = trunc.replace(/\w+$/, '');
			        trunc += '... <a href="sites.asp?region='+region+'&locid='+locid+'" class="sitepageLinkOff" onmouseover="this.className=\'sitepageLinkOn\';" onmouseout="this.className=\'sitepageLinkOff\';" style="text-decoration:underline;">More info&nbsp;</a><img src="http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_arrow-sm.gif">';
			        return trunc;
		        }
	        }
        }
        function truncateLoc(txt)
        {
	        var len = 32;
	        var p = txt;
	        if (p)
	        {
		        var trunc = p;
		        if (trunc.length > len)
		        {
			        trunc = trunc.substring(0, len);
			        trunc = trunc.replace(/\w+$/, '');
			        trunc += "...";
			        return trunc;
		        }
	        }
        }        
        function addEvent(obj,evType,fn){if(obj.addEventListener){obj.addEventListener(evType,fn,false);return true;}else if(obj.attachEvent){var r=obj.attachEvent("on"+evType,fn);return r;}else{return false;}}

        addEvent(window, 'load', GetMap);
	</script>
	<style type="text/css">
    .MSVE_ZoomBar {
        background-color:#5a7a9b;
        color:ffffff;
        filter:alpha(opacity=10);
        opacity:0.1;
    }
    </style>
    <div id="mapContainer" style="width:420px; height:216px; border:solid 1px #cccccc;">
	    <div id="myMap" style="position:relative; width:418px; height:214px;"></div>
    </div>
    <br />
    <table border="0" cellpadding="0" cellspacing="0" width="418">
        <tr>
             <td>
                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                        <td valign="top" style="padding-left:8px; width:17px;">
                            <img src="http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_map_arrow_nomap.gif" border="0" width="16" height="17" />
                        </td>
                        <td align="left" valign="top" style="color:#264466; font-size:14px; font-family:arial, helvetica, sans serif; font-weight:normal; padding-left:6px; padding-top:0px; vertical-align:top;">
                            <p style="margin-top:0px; margin-left:0px;">Popular World Heritage sites:</p>
                        </td>            
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td id="linkHolder" align="left" style="padding-left:10px; width:418px;" class="linkHolder">
            </td>
        </tr>
        <tr>
            <td style="font-family:arial, helvetica, sans serif; font-size: 12px; padding-top:10px; padding-left:10px;">
                <img src="http://media.expedia.com/media/content/expus/graphics/promos/deals/wh_landing_arrow.gif" align="absmiddle" style="border:none;padding-right:2px;"> <a href="find.asp" onmouseover="this.className='sitepageLinkOn'" onmouseout="this.className='sitepageLinkOff'" class="sitepageLink">Find a World Heritage site</a>
            </td>
        </tr>
    </table>
<script language="javascript" type="text/javascript">
    if (navigator.userAgent.indexOf("MSIE 7") != -1)
    {
        document.getElementById("myMap").style.cssText = "position:relative;width:418px;height:214px;margin:1px 8px 8px 1px;";
    }
    if (navigator.userAgent.indexOf("Firefox") != -1)
    {
        document.getElementById("myMap").style.cssText = "position:relative;width:418px;height:214px;margin:1px 8px 8px 1px;";
    }
</script>