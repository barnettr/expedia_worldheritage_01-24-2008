function SetMapStyle()
{
    if (document.styleForm.styleType[0].checked)
    {
        var style = map.GetMapStyle();
        if (style == "a")
        {
        }
        else
        {
            map.SetMapStyle("a");
        }
    }
    else if (document.styleForm.styleType[1].checked)
    {
        var style = map.GetMapStyle();
        if (style == "r")
        {
        }
        else
        {
            map.SetMapStyle("r");
        }
    }
    else if (document.styleForm.styleType[2].checked)
    {
        var style = map.GetMapStyle();
        if (style == "h")
        {
        }
        else
        {
            map.SetMapStyle("h");
        }
    }
}

