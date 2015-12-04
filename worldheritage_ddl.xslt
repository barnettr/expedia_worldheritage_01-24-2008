<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:fn="http://www.w3.org/2005/02/xpath-functions" xmlns:xdt="http://www.w3.org/2005/02/xpath-datatypes">
	<xsl:output method="html" indent="yes"/>
	<xsl:key name="distinct-state" match="state" use="." />
	<xsl:key name="group-state" match="row" use="state" />
	<xsl:param name="region" />
	<xsl:param name="state" />
    <xsl:param name="pview" />
    <xsl:param name="url" />	
	<xsl:template match="worldheritage">
	<select style="width:150px;font-size:11px;">
	<xsl:choose>
		<xsl:when test="$region = ''">
			<xsl:attribute name="disabled">disabled</xsl:attribute>
		</xsl:when>
		<xsl:otherwise>
			<xsl:attribute name="onChange">window.location=this.options[this.selectedIndex].value;</xsl:attribute>
		</xsl:otherwise>
	</xsl:choose>
		<option><xsl:if test="row/region[1]"><xsl:attribute name="selected">selected</xsl:attribute></xsl:if><xsl:attribute name="value"><xsl:value-of select="$url" />?region=<xsl:value-of select="translate(row/region[1],' ','_')" />&amp;pview=<xsl:value-of select="$pview" /></xsl:attribute>All Locations&#160;(<xsl:value-of select="count(row)" />)</option>
		<xsl:for-each select="row">
			<xsl:sort order="ascending" select="state" />
			<xsl:variable name="uniqueState" select="state[count(.|key('distinct-state',.)[1]) = last()]" />
			<xsl:variable name="totalState" select="count(.|key('group-state',state))" />
			<xsl:if test="$uniqueState">
			<xsl:for-each select="key('group-state', state)"></xsl:for-each>
			<option><xsl:if test="$uniqueState = $state"><xsl:attribute name="selected">selected</xsl:attribute></xsl:if> <xsl:attribute name="value"><xsl:value-of select="$url" />?region=<xsl:value-of select="region[1]" />&amp;state=<xsl:value-of select="$uniqueState" />&amp;pview=<xsl:value-of select="$pview" /></xsl:attribute><xsl:value-of select="$uniqueState" />&#160;(<xsl:value-of select="$totalState" />)</option></xsl:if>
		</xsl:for-each>
		</select>
</xsl:template> 
</xsl:stylesheet>
