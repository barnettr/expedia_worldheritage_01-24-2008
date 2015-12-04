<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xs="http://www.w3.org/2001/XMLSchema">
  <xsl:output omit-xml-declaration="no" method="html" indent="yes"/>
  <xsl:key name="distinct-reg" match="region" use="." />  
  <xsl:key name="distinct-state" match="state" use="." />
  <xsl:key name="group-reg" match="row" use="region" />  
  <xsl:key name="group-state" match="row" use="state" />
  <xsl:param name="region" />
  <xsl:param name="state" />
  <xsl:param name="page" />
  <xsl:param name="pview" />
  <xsl:param name="maxPages" />
  <xsl:param name="SortCol" />
  <xsl:param name="SortDir" />
  <xsl:param name="SortType" />
  <xsl:template match="worldheritage">
    <style type="text/css">
.gridIcon
{
	background-image: url(grid_button.gif);
	background-repeat: no-repeat;
	color: #ffffff;
	cursor: hand;
	font-family: Trebuchet MS,arial,sans-serif;
	font-size: 11px;
	font-weight: bold;
	height: 20px;
	padding-top: 0px;
	width: 24px;
}
.gridIconOff
{
	background-image: url(grid_button_selected.gif);
	background-repeat: no-repeat;
	color: #cc6600;
	cursor: hand;
	font-family: Trebuchet MS,arial,sans-serif;
	font-size: 11px;
	font-weight: bold;
	height: 20px;
	padding-top: 0px;
	width: 24px;
}
html>body .gridIcon
{
}
#pagination
{
	background-color: #ffffff;
	font-family: Trebuchet MS,arial,sans-serif;
	font-size: 11px;
	padding-bottom: 4px;
	padding-left: 4px;
	padding-top: 4px;
	text-align: left;
	vertical-align: middle;
}
#pagination select
{
	font-size: 10px;
	margin: 0;
	padding: 0;
	vertical-align: middle;
}
#pagination img
{
	margin: 0px;
	padding-bottom: 4px;
	padding-top: 0px;
	vertical-align: middle;
}
#pagination a, #pagination .nopage
{
	border: 1px solid #fff;
	color: #336699;
	font-weight: bold;
	margin: 0;
	padding: 0 4px;
	text-decoration: none;
}
#pagination a:hover, #pagination a:active
{
	background-image: url(http://media.expedia.com/media/content/expus/graphics/promos/deals/1007_paradise_pagination_box_1x19.gif);
	border: 1px solid #666;
	color: #ff6600;
	font-family: Trebuchet MS, arial,sans-serif;
	font-size: 11px;
	text-align: center;
	text-decoration: underline;
}
#pagination .current
{
	background-image: url(http://media.expedia.com/media/content/expus/graphics/promos/deals/1007_paradise_pagination_box_1x19.gif);
	border: 1px solid #666;
	color: #000;
	font-family: Trebuchet MS, arial,sans-serif;
	font-size: 11px;
	font-weight: bold;
	margin: 2px;
	padding: 0 4px;
	text-align: center;
}
td.paging
{
	text-align: left;
}
td.pagingBottom
{
	padding: 20px 0 16px 0;
}
.sitenumber
{
	background-color: #4985c7;
	border: 1px solid #EEE;
	color: #FFF;
	font-weight: bold;
	margin: 2px;
	padding: 1px 3px 1px 3px;
}
.rowHover
{
	background-color: #e7eff7;
	cursor: pointer;
}
.siteRow
{
	border-top: 1px solid #F5F5F5;
	color: #333;
	font-family: Trebuchet MS, arial,sans-serIf;
	font-size: 11px;
	margin: 0;
	padding: 0;
}
.four
{
	border-bottom: 1px solid #;
	color: #336699;
	font-family: Trebuchet MS, arial, helvetica, sans serif;
	font-size: 12px;
	font-weight: bold;
	padding-left: 4px;
	padding-top: 3px;
}
.cellOne
{
	border-bottom: 1px dotted #e0dfdf;
	color: #ffffff;
	font-family: Trebuchet MS, arial, helvetica, sans serif;
	font-size: 11px;
	font-weight: bold;
	padding-left: 4px;
}
}
    </style>
    <!-- LOOP THROUGH ALL ROWS -->
  	
  <script type="text/javascript">
	function setPageView(type,opt) {
		var url;
		var qs = new String(window.location);
		var pview = /pview\=\d+/i;
		var page = /page\=\d+/i;
		if (qs.split('?')[1] == undefined) {
			url = (type=='pview')?qs+'?pview='+opt:qs+'?page='+opt;
		} else {
			if (type=='pview') {
				url = (qs.match(pview)!=null) ? qs.replace(pview,'pview='+opt) : qs+'&amp;pview='+opt;
				url = url.replace(page,'page=1');				
			} else {
				url = (qs.match(page)!=null) ? qs.replace(page,'page='+opt) : qs+'&amp;page='+opt;
			}
		}
		document.location=url;
	}
	
	function doSort(col) {
		var url;
		var qs = new String(window.location);
		var sortCol = /sortCol\=[a-z0-9]+/i;
		var sortDir = /sortDir\=[a-z]+/i;
		var page = /page\=\d+/i;		
		if (qs.split('?')[1] == undefined) {
			url = qs+'?sortCol='+col+'&amp;sortDir=descending';
		} else {
			if (qs.match(sortCol)!=null) {
				var curSortCol = qs.match(sortCol);
				var curSortDir = qs.match(sortDir);
				var selSortCol = 'sortCol='+col;
				var selSortDir = (curSortDir=='sortDir=ascending')?'sortDir=descending':'sortDir=ascending';
				url = (curSortCol==selSortCol)?qs.replace(sortDir,selSortDir):qs.replace(sortCol,selSortCol).replace(sortDir,'sortDir=ascending');
			} else {
				url = qs+'&amp;sortCol='+col+'&amp;sortDir=descending';
			}
		}
		url = url.replace(page,'page=1');
		document.location=url;
	}
	
	var CurrentState;
	</script>
  <table id="grid" border="0" cellpadding="0" cellspacing="0">
    <xsl:for-each select="row">
      <xsl:variable name="uniqueReg" select="region[count(.|key('distinct-reg',.)[1]) = last()]" />
      <xsl:variable name="uniqueState" select="state[count(.|key('distinct-state',.)[1]) = last()]" />
      <!-- CHECK IF FILTERING A SPECIFIC STATE -->
      <xsl:if test="$uniqueState = state and $state = state">
	    <tr><td colspan="4"><xsl:call-template name="paging" /></td></tr>
        <!-- LOOP THROUGH ALL state -->
        <!--<tr>
      <th bgcolor="#3e6d9f"> </th>
		  <th>Photo</th>-->
        <!--
        <th bgcolor="#3e6d9f" style="font-family: arial, helvetica, sans serif; font-size: 14px; color: #ffffff; font-weight: bold;">Site Name</th>
		  <th align="right" bgcolor="#3e6d9f" style="font-family: arial, helvetica, sans serif; font-size: 14px; color: #ffffff; font-weight: bold; padding-right:3px;" onclick="doSort('state')">Destination</th>
        </tr>-->
        <xsl:for-each select="key('group-state', state)">
			<xsl:sort order="{$SortDir}" select="*[name()=$SortCol]" data-type="{$SortType}" />
			<xsl:if test="(position() >= 1 + ($page - 1) * $pview) and (position() &lt; (1 + $page * $pview))">
            	<xsl:call-template name="grid" />  
			</xsl:if>        
        </xsl:for-each>
      </xsl:if>
      <!-- CHECK IF FILTERING A SPECIFIC REGION -->
      <xsl:if test="$uniqueReg = region and $region = region and $state = ''">
	  <tr><td colspan="4"><xsl:call-template name="paging" /></td></tr>
        <!-- LOOP THROUGH ALL GROUPS -->
    <!--<tr>-->
      <!--<th bgcolor="#3e6d9f"> </th>-->
      <!--<th>Photo</th>-->
      <!--<th bgcolor="#3e6d9f" style="font-family: arial, helvetica, sans serif; font-size: 14px; color: #ffffff; font-weight: bold;">Site Name</th>
      <th align="right" bgcolor="#3e6d9f" style="font-family: arial, helvetica, sans serif; font-size: 14px; color: #ffffff; font-weight: bold; padding-right:3px;" onclick="doSort('state')">Destination</th>-->
        <!--</tr>-->
        <xsl:for-each select="key('group-reg', region)">
			<xsl:sort order="{$SortDir}" select="*[name()=$SortCol]" data-type="{$SortType}" />
			<xsl:if test="(position() >= 1 + ($page - 1) * $pview) and (position() &lt; (1 + $page * $pview))">
				<xsl:call-template name="grid" />
			</xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:for-each>
	</table>
  </xsl:template>
  <!-- BEGIN PRICE GRID -->
  
  <xsl:template name="grid">
    <tr bgcolor="#f8f8f8" style="line-height:18px;">
      <xsl:attribute name="id">row_<xsl:value-of select="position()" /></xsl:attribute>
        <xsl:attribute name="onmouseover">GridShowInfoBox(getPinID(<xsl:value-of select="position()" />));showRow('<xsl:value-of select="position()" />','row');</xsl:attribute>
        <xsl:attribute name="onmouseout">GridHideInfoBox(getPinID(<xsl:value-of select="position()" />));hideRow('<xsl:value-of select="position()" />','row');</xsl:attribute>
      <td width="24"  align="center" class="cellOne">
        <script type="text/javascript">


          //Create an object with these properties then add to array
          //BuildPushpin(id,site,suitcase,flag,desc,img,lat,lng)
          var pin = new Object();
          pin.locid = "<xsl:value-of select="@id" />";
          pin.region = "<xsl:value-of select="region" />";
          pin.id = "<xsl:value-of select="position()" />";
          pin.site = "<xsl:value-of select="site" />";
          pin.suitcase = "suitcase.gif";
          pin.flag = "http://whc.unesco.org/uploads/states/small/<xsl:value-of select="code" />.gif";
          pin.desc = "<xsl:value-of select="description" />";
          pin.img = "http://whc.unesco.org/uploads/sites/site_<xsl:value-of select="@id" />.jpg";
          pin.lat = "<xsl:value-of select="@lat" />";
          pin.lng = "<xsl:value-of select="@long" />";
          arrPin[pinid] = pin;
          pinid++;
        </script>
        <div class="gridIcon" onmouseover="this.className='gridIconOff';"  onmouseout="this.className='gridIcon';"><xsl:attribute name="id">gridIcon_<xsl:value-of select="position()" /></xsl:attribute>
          <xsl:value-of select="position()" />
        </div>
      </td>
      <!--<td><img><xsl:attribute name="src">http://whc.unesco.org/uploads/sites/site_<xsl:value-of select="@id" />.jpg</xsl:attribute></img></td>-->
      <td onmouseover="style.color='#4985c7';" onmouseout="style.color='#4985c7';" style="padding-bottom:4px;padding-top:3px;padding-left:5px;padding-right:4px;font-family: Trebuchet MS, arial, helvetica, sans serif;font-size: 11px;color: #4985c7;font-weight: bold;cursor:pointer;text-decoration:none;border-bottom: 1px dotted #e0dfdf;">
        <span><xsl:attribute name="title"><xsl:value-of select="site" /></xsl:attribute><xsl:value-of select="substring(site,0,50)" disable-output-escaping="yes" /></span>
        <br />
        <xsl:choose>
          <xsl:when test="location != ''">
            <span style="color:#666666;"><xsl:value-of select="state" />, <xsl:value-of select="region" />
        </span>
          </xsl:when>
          <xsl:otherwise>
            Unknown
          </xsl:otherwise>
        </xsl:choose>
        <br />
      </td>
    </tr>


  </xsl:template>
  
  <!-- END PRICE GRID -->
  <!-- START PAGING -->
  <xsl:template name="paging">
    <xsl:variable name="TotalItems">
    <xsl:choose>
      <xsl:when test="$state = '' and $region != ''">
        <xsl:value-of select="count(.|key('group-reg', region))" />
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="count(.|key('group-state', state))" />
      </xsl:otherwise>
    </xsl:choose>
    </xsl:variable>
    <xsl:variable name="Pages" select="ceiling($TotalItems div $pview)" />
    <!-- select first element of each page -->
   <!--Results per page: 
      <select id="ddlPageSize" onchange="setPageView('pview',this.options[this.selectedIndex].value);">
        <xsl:choose>
          <xsl:when test="$pview='10'">
            <option value="10" selected="selected">10</option>
            <option value="25">25</option>
            <option value="50">50</option>
            <option value="99">99</option>
          </xsl:when>
          <xsl:when test="$pview='25'">
            <option value="10">10</option>
            <option value="25" selected="selected">25</option>
            <option value="50">50</option>
            <option value="99">99</option>
          </xsl:when>
          <xsl:when test="$pview='50'">
            <option value="10">10</option>
            <option value="25">25</option>
            <option value="50" selected="selected">50</option>
            <option value="99">99</option>
          </xsl:when>
          <xsl:when test="$pview='99'">
            <option value="10">10</option>
            <option value="25">25</option>
            <option value="50">50</option>
            <option value="99" selected="selected">99</option>
          </xsl:when>
        </xsl:choose>
      </select>--> <div id="pagination" style="border-bottom: 1px dotted #e0dfdf; margin-top:2px;">
        <!--<img src="http://media.expedia.com/media/content/expus/graphics/promos/cruise/tab_lineSpacer.gif" width="1" height="14" border="0" />--> Page
	  <xsl:choose>
		  <xsl:when test="$page>1">
			  <a>
				  <xsl:attribute name="href">javascript:setPageView('page','<xsl:value-of select="$page - 1" />');</xsl:attribute>
				  <xsl:attribute name="alt">Previous Page</xsl:attribute>
				  <xsl:attribute name="title">Previous Page</xsl:attribute>
				  <img src="http://media.expedia.com/media/content/expus/graphics/promos/hotel/pagination_arrowlt_on_10x14.gif" alt="Last" border="0"  />
			  </a>
		  </xsl:when>
		  <xsl:otherwise>
			  <img>
			  	<xsl:attribute name="src">http://media.expedia.com/media/content/expus/graphics/promos/hotel/pagination_arrowlt_off_10x14.gif</xsl:attribute>
				<xsl:attribute name="class">nopage</xsl:attribute>
			  </img>
		  </xsl:otherwise>
	  </xsl:choose>
	  <xsl:choose>
        <xsl:when test="$state = '' and $region != ''">
          <xsl:for-each select="key('group-reg', region)[((position()-1) mod $pview = 0)]">
            <xsl:call-template name="pages" /></xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="key('group-state', state)[((position()-1) mod $pview = 0)]">
            <xsl:call-template name="pages" /></xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
	  <xsl:choose>
		  <xsl:when test="$page &lt; ceiling($TotalItems div $pview)">
			  <a>
				  <xsl:attribute name="href">javascript:setPageView('page','<xsl:value-of select="$page + 1" />');</xsl:attribute>
				  <xsl:attribute name="alt">Next Page</xsl:attribute>
				  <xsl:attribute name="title">Next Page</xsl:attribute>
				  <img src="http://media.expedia.com/media/content/expus/graphics/promos/hotel/pagination_arrowrt_on_10x14.gif" alt="Last" border="0"  />
			  </a>
		  </xsl:when>
		  <xsl:otherwise>
			  <img>
			  	<xsl:attribute name="src">http://media.expedia.com/media/content/expus/graphics/promos/hotel/pagination_arrowrt_off_10x14.gif</xsl:attribute>
				<xsl:attribute name="class">nopage</xsl:attribute>
			  </img>
		  </xsl:otherwise>
	  </xsl:choose>	  
    </div>
  </xsl:template>
  <!-- END PAGING -->
  <!-- BEGIN PAGES -->
  <xsl:template name="pages">
    <!-- display the appropriate portion of page links; disable link to active page -->
    <xsl:choose>	
      <xsl:when test="(position() > ($page - ceiling($maxPages div 2)) or position() > (last() - $maxPages)) and ((position() &lt; $page + $maxPages div 2) or (position() &lt; 1 + $maxPages))">
        <xsl:if test="position()=$page">
          <span class="current"><xsl:value-of select="position()"/></span>
        </xsl:if>
        <xsl:if test="position()!=$page">
          <a>
          <xsl:attribute name="href">javascript:setPageView('page','<xsl:value-of select="position()" />');</xsl:attribute>
          <xsl:attribute name="alt">Page&#160;<xsl:value-of select="position()" /></xsl:attribute>
          <xsl:attribute name="title">Page&#160;<xsl:value-of select="position()" /></xsl:attribute>
          <xsl:value-of select="position()"/></a>
        </xsl:if>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!-- END PAGES -->  
</xsl:stylesheet>
