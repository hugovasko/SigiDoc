<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet exclude-result-prefixes="#all"
                version="2.0"
                xmlns:tei="http://www.tei-c.org/ns/1.0"
                xmlns:fn="http://www.w3.org/2005/xpath-functions"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

  <!-- This XSLT transforms a set of EpiDoc documents into a Solr
       index document representing an index of symbols in those
       documents. -->

  <xsl:import href="epidoc-index-utils.xsl" />

  <xsl:param name="index_type" />
  <xsl:param name="subdirectory" />
  <xsl:variable name="dignities" select="doc('../../content/xml/authority/dignities.xml')"/>

  <xsl:variable name="map_points">
    <xsl:text>[</xsl:text>
    <xsl:for-each select="collection('../../content/xml/epidoc/?select=*.xml;recurse=yes')//tei:teiHeader[matches(normalize-space(descendant::tei:geo), '\d{1,2}(\.\d+){0,1},\s+?\d{1,2}(\.\d+){0,1}')]">
      <xsl:text>{</xsl:text>
      <xsl:text>"title": "</xsl:text><xsl:value-of select="normalize-space(descendant::tei:title/tei:seg[1])"/><xsl:text>",</xsl:text>
      <xsl:text>"coordinates": "</xsl:text><xsl:value-of select="normalize-space(descendant::tei:geo[1])"/><xsl:text>",</xsl:text>
      <xsl:text>"findspot": "</xsl:text><xsl:value-of select="normalize-space(descendant::tei:findspotAccuracy[1])"/><xsl:text>",</xsl:text>
      <xsl:text>"date": "</xsl:text><xsl:value-of select="normalize-space(descendant::tei:origDate[1]/tei:seg[1])"/><xsl:text>",</xsl:text>
      <xsl:text>"filename": "</xsl:text><xsl:value-of select="tokenize(base-uri(ancestor::tei:TEI[1]),'/')[last()]"/><xsl:text>"</xsl:text>
      <xsl:text>}</xsl:text>
      <xsl:if test="position()!=last()"><xsl:text>,</xsl:text></xsl:if>
    </xsl:for-each>
    <xsl:text>]</xsl:text>
  </xsl:variable>

  <xsl:template match="/">
    <add>
      <xsl:result-document href="webapps/ROOT/assets/leafletMap/map_points.json" method="text">
        <xsl:value-of select="$map_points" />
      </xsl:result-document>
      <xsl:for-each-group select="//tei:rs[@type='dignity'][@ref][ancestor::tei:div/@type='textpart']" group-by="@ref">
        <doc>
          <field name="document_type">
            <xsl:value-of select="$subdirectory" />
            <xsl:text>_</xsl:text>
            <xsl:value-of select="$index_type" />
            <xsl:text>_index</xsl:text>
          </field>
          <xsl:call-template name="field_file_path" />
          <field name="index_item_name">
              <xsl:variable name="ref-id" select="substring-after(@ref,'#')"/>
              <xsl:value-of select="$dignities//tei:item[@xml:id = $ref-id]//tei:term[@xml:lang = 'grc' or @xml:lang = 'la']" />
          </field>
          <xsl:apply-templates select="current-group()" />
        </doc>
      </xsl:for-each-group>
    </add>
  </xsl:template>

  <xsl:template match="tei:rs">
    <xsl:call-template name="field_index_instance_location" />
  </xsl:template>

</xsl:stylesheet>
