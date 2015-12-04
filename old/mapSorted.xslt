<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:output method="xml" omit-xml-declaration="yes" indent="yes" />
  <xsl:param name="SortCol" />
  <xsl:param name="SortDir" />
  <xsl:param name="SortType" />
  <xsl:template match="worldheritage">
    <worldheritage promo="UNESCO">
      <xsl:for-each select="row">
        <xsl:sort order="{$SortDir}" select="*[name()=$SortCol]" data-type="{$SortType}" />
        <xsl:copy-of select="."/>
      </xsl:for-each>
    </worldheritage>
  </xsl:template>
</xsl:stylesheet>