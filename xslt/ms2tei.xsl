<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:sheet="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns="http://www.tei-c.org/ns/1.0"
  exclude-result-prefixes="#all"
  version="2.0">
  
  <xsl:output indent="yes" />
  <xsl:param name="heading" select="false()" />
  <xsl:param name="title" />
  <xsl:param name="filename" />
  
  <xsl:template name="extract">
    <xsl:apply-templates select="/" />
  </xsl:template>
  
  <xsl:template match="/">
    <xsl:apply-templates />
  </xsl:template>
  
  <xsl:template match="pkg:package">
    <xsl:apply-templates select="pkg:part[contains(@pkg:name, 'workbook.xml')]//sheet:sheet" />
  </xsl:template>
  
  <xsl:template match="sheet:sheet">
    <xsl:variable name="sheet">
      <xsl:variable name="id" select="@r:id" />
      <xsl:value-of select="substring-after(//rel:Relationship[@Id = $id]/@Target, '/')"/>
    </xsl:variable>
    <xsl:result-document href="{@name}.xml">
      <TEI version="5.0">
        <teiHeader>
          <fileDesc>
            <titleStmt>
              <title>
                <xsl:value-of select="$title"/>
              </title>
            </titleStmt>
            <publicationStmt>
              <p>No information available</p>
            </publicationStmt>
            <sourceDesc>
              <p>Created from <xsl:value-of select="$filename" /> by table2TEI, ms2TEI.xsl on <xsl:value-of select="current-dateTime()"/></p>
            </sourceDesc>
          </fileDesc>
        </teiHeader>
        <text>
          <body>
            <xsl:apply-templates select="//pkg:part[ends-with(@pkg:name, $sheet)]//sheet:sheetData" />
          </body>
        </text>
      </TEI>
    </xsl:result-document>
  </xsl:template>
  
  <xsl:template match="sheet:sheetData">
    <table cols="{count(preceding-sibling::sheet:cols/sheet:col)}" rows="{count(sheet:row[descendant::sheet:v])}">
      <xsl:apply-templates select="sheet:row[descendant::sheet:v]">
        <xsl:with-param name="cols" select="count(preceding-sibling::sheet:cols/sheet:col)" />
      </xsl:apply-templates>
    </table>
  </xsl:template>
  
  <xsl:template match="sheet:row[not(sheet:c)]" />
  <xsl:template match="sheet:row[sheet:c]">
    <xsl:param name="cols" />
    <row>
      <xsl:variable name="free"
          select="ancestor::sheet:sheetData/preceding-sibling::sheet:sheetViews/sheet:sheetView/sheet:pane[@state='frozen']/@topLeftCell" />
      <xsl:variable name="row" select="."/>
      <xsl:if test="(string-length($free) &gt; 0 and following-sibling::sheet:row/sheet:c[@r=$free])
        or (not(preceding-sibling::sheet:row) and $heading)">
        <xsl:attribute name="role">label</xsl:attribute>
      </xsl:if>
      <xsl:for-each select="(1 to $cols)">
        <xsl:variable name="col">
          <xsl:number value="." format="A" />
          <xsl:value-of select="$row/@r"/>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="$row/sheet:c[@r = $col]">
            <xsl:apply-templates select="$row/sheet:c[@r = $col || @row]" />
          </xsl:when>
          <xsl:otherwise>
            <cell />
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </row>
  </xsl:template>
  
  <xsl:template match="sheet:c">
    <cell>
      <xsl:choose>
        <xsl:when test="@t='s'">
          <xsl:variable name="num" select="sheet:v + 1"/>
          <xsl:apply-templates select="//sheet:sst/sheet:si[$num]/sheet:t"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:apply-templates select="sheet:v" />
        </xsl:otherwise>
      </xsl:choose>
    </cell>
  </xsl:template>
</xsl:stylesheet>