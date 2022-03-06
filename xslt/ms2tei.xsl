<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
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
  <xsl:param name="server" />
  
  <xsl:template name="extract">
    <xsl:apply-templates select="/" />
  </xsl:template>
  
  <xsl:template match="/">
    <xsl:apply-templates />
  </xsl:template>
  
  <xsl:template match="pkg:package">
    <xsl:apply-templates select="pkg:part[@pkg:name =  '/xl/workbook.xml']//sheet:sheet" />
  </xsl:template>
  
  <xsl:template match="sheet:sheet">
    <xsl:variable name="id" select="@r:id" />
    <xsl:variable name="sheet"
      select="//pkg:part[@pkg:name='/xl/_rels/workbook.xml.rels']//rel:Relationship[@Id = $id]/@Target"/>
    
    <xsl:apply-templates select="//pkg:part[@pkg:name = '/xl/' || $sheet]">
      <xsl:with-param name="name" select="@name" tunnel="yes" />
      <xsl:with-param name="pkgName" select="substring-after($sheet, '/')" tunnel="yes" />
    </xsl:apply-templates>
  </xsl:template>
  
  <xsl:template match="pkg:part">
    <xsl:param name="name" tunnel="yes" />
    
    <xsl:choose>
      <xsl:when test="$server != ''">
        <xsl:call-template name="result" />
      </xsl:when>
      <xsl:otherwise>
        <xsl:result-document href="{$name}.xml">
          <xsl:call-template name="result" />
        </xsl:result-document>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="result">
    <xsl:param name="name" tunnel="yes" />
    
    <TEI version="5.0">
      <teiHeader>
        <fileDesc>
          <titleStmt>
            <title>
              <xsl:value-of select="$name"/>
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
          <xsl:apply-templates select="descendant::sheet:sheetData" />
        </body>
      </text>
    </TEI>
  </xsl:template>
  
  <xsl:template match="sheet:sheetData">
    <xsl:param name="pkgName" tunnel="yes" />
    <xsl:variable name="cols" select="count(preceding-sibling::sheet:cols/sheet:col)" />
    
    <table cols="{$cols}" rows="{count(sheet:row[descendant::sheet:v])}">
      <xsl:apply-templates select="sheet:row[descendant::sheet:v]">
        <xsl:with-param name="cols" select="$cols" tunnel="yes" />
        <xsl:with-param name="rels" select="'/xl/worksheets/_rels/' || $pkgName || '.rels'" tunnel="yes" />
      </xsl:apply-templates>
    </table>
  </xsl:template>
  
  <xsl:template match="sheet:row[not(sheet:c)]" />
  <xsl:template match="sheet:row[sheet:c]">
    <xsl:param name="cols" tunnel="yes" />
    
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
    <xsl:param name="rels" tunnel="yes" />
    <xsl:variable name="r" select="@r" />
    <xsl:variable name="stylePosition" select="number(@s) + 1" />
    
    <cell>
      <xsl:apply-templates select="//sheet:cellXfs/sheet:xf[position() = $stylePosition]" />
      <xsl:choose>
        <xsl:when test="ancestor::sheet:worksheet/sheet:hyperlinks/sheet:hyperlink[@ref = $r]">
          <xsl:variable name="hyperlink" select="ancestor::sheet:worksheet/sheet:hyperlinks/sheet:hyperlink[@ref = $r]"/>
          <xsl:variable name="relId" select="$hyperlink/@r:id" />
          <xsl:variable name="rel" select="//pkg:part[@pkg:name = $rels]//rel:Relationship[@Id = $relId]"/>
          
          <ref type="link" target="{$rel/@Target}">
            <xsl:value-of select="$hyperlink/@display"/>
          </ref>
        </xsl:when>
        <xsl:when test="@t='s'">
          <xsl:variable name="num" select="sheet:v + 1"/>
          <xsl:apply-templates select="//sheet:sst/sheet:si[$num]/*"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:apply-templates select="sheet:v" />
        </xsl:otherwise>
      </xsl:choose>
    </cell>
  </xsl:template>
  
  <xsl:template match="sheet:r">
    <xsl:apply-templates select="sheet:t" />
  </xsl:template>
  
  <xsl:template match="sheet:r[sheet:rPr]">
    <hi>
      <xsl:attribute name="style">
        <xsl:apply-templates select="sheet:rPr/*" />
      </xsl:attribute>
      <xsl:apply-templates select="sheet:t" />
    </hi>
  </xsl:template>
  
  <xsl:template match="sheet:i[not(@val) or @val = '1']">
    <xsl:text>font-variant: italic;</xsl:text>
  </xsl:template>
  
  <xsl:template match="sheet:sz">
    <xsl:value-of select="'font-size: ' || @val || ';'" />
  </xsl:template>
  
  <xsl:template match="sheet:rFont">
    <xsl:value-of select="'font-family: ' || @val || ';'" />
  </xsl:template>
  
  <xd:doc>
    <xd:desc>Evaluate a style definition</xd:desc>
  </xd:doc>
  <!-- TODO: This needs further expansion to correctly incorporate fillId, boerderId, xfId and evaluate the apply* attributes -->
  <xsl:template match="sheet:xf">
    <xsl:variable name="fontPosition" select="number(@fontId) + 1" />
    
    <xsl:attribute name="style">
      <xsl:apply-templates select="ancestor::sheet:styleSheet/sheet:fonts/sheet:font[position() eq $fontPosition]/*" />
    </xsl:attribute>
  </xsl:template>
</xsl:stylesheet>
