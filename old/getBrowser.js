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

function getBrowser()
{
    if (navigator.userAgent.indexOf("MSIE") != -1)
    {
        g_isIE = true;
        getCustomPushpin();
        
        document.getElementById("logo").style.cssText = "display:none; z-index:0; visibility:hidden;";
        document.getElementById("inprogress_img").style.cssText = "display:block; z-index:30; position:absolute; left:850px; top:280px; width:100px; height:20px; background-color:#d7e8e6; border:solid 2px #a31818; font-weight:bold; color:#666666; font-family:Arial; font-size:11px; vertical-align:middle;";
        if (navigator.userAgent.indexOf("MSIE 7") != -1)
        {
            g_isIE7 = true;
        }
    }
    else if (navigator.userAgent.indexOf("Firefox") != -1)
    {
        g_isFirefox = true;
        loadXML();
        document.getElementById("logo").style.cssText = "display:inline; z-index:30; position:absolute; bottom:2px; left:8px";
        document.getElementById("showhide").style.cssText = "display:block; padding:1px 1px 1px 1px; margin:0px; border:1px solid #e6e7e7; font-weight:bold; background-color:#f4f9f9; color:#000000; text-decoration:none; font-family:Arial; font-size:8px; position:absolute; left:37px; bottom:442px; z-index:30;";
        document.getElementById("inprogress_img").style.cssText = "display:block; z-index:30; position:absolute; left:250px; bottom:300px; width:100px; height:20px; background-color:#d7e8e6; border:solid 2px #a31818; font-weight:bold; color:#666666; font-family:Arial; font-size:11px; vertical-align:top;";
        document.getElementById("p_holder").style.cssText = "position:relative; bottom:8px;";
        document.getElementById("loader").style.cssText = "position:relative; bottom:1px;";
    }
}