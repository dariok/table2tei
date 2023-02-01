<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:archive="http://expath.org/ns/archive"
  xmlns:file="http://expath.org/ns/file"
  xmlns:zip="http://www.expath.org/mod/zip"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:sheet="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:err="http://www.w3.org/2005/xqt-errors"
  xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships"
  exclude-result-prefixes="#all"
  version="3.0">
  
  <xsl:output indent="1" />
  <xsl:param name="filename" />
  
  <xsl:template match="/">
    <xsl:if test="not(function-available('file:read-binary')) and not(function-available('zip:xml-entry'))">
      <xsl:message terminate="yes">!!
        Neither file:read-binary nor zip:xml-entry functions are available. Word document
        cannot be extracted with this XSLT processor. zip:xml-entry is available in Saxon
        ⩽ 9.5.1.1, file:read-binary requires Saxon PE or EE ⩾ 9.6.
      </xsl:message>
    </xsl:if>
    
    <xsl:variable name="zip">
      <xsl:sequence select="file:read-binary($filename)" use-when="function-available('archive:extract-text')" />
      <xsl:value-of select="xs:anyURI($filename)" use-when="function-available('zip:xml-entry')" />
    </xsl:variable>
    
    <xsl:variable name="relations">
      <xsl:sequence select="parse-xml(archive:extract-text($zip, 'xl/_rels/workbook.xml.rels'))" use-when="function-available('archive:extract-text')"/>
      <xsl:sequence select="zip:xml-entry($zip, 'xl/_rels/workbook.xml.rels')" use-when="function-available('zip:xml-entry')"/>
    </xsl:variable>
     
     <xsl:variable name="parts" select="('xl/workbook.xml', 'xl/_rels/workbook.xml.rels')" />
    
    <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <xsl:for-each select="$parts">
        <xsl:try>
          <pkg:part>
            <xsl:attribute name="pkg:name" select="'/' || ." />
            <pkg:xmlData>
              <xsl:sequence select="parse-xml(archive:extract-text($zip, .))" use-when="function-available('archive:extract-text')"/>
              <xsl:sequence select="zip:xml-entry($zip, .)" use-when="function-available('zip:xml-entry')"/>
            </pkg:xmlData>
          </pkg:part>
          <xsl:catch />
        </xsl:try>
      </xsl:for-each>
      
      <xsl:for-each select="$relations//r:Relationship">
        <xsl:variable name="path" select="'xl/' || @Target" />
        <pkg:part>
          <xsl:attribute name="pkg:name" select="'/' || $path" />
          <pkg:xmlData>
            <xsl:sequence select="parse-xml(archive:extract-text($zip, $path))" use-when="function-available('archive:extract-text')"/>
            <xsl:sequence select="zip:xml-entry($zip, $path)" use-when="function-available('zip:xml-entry')"/>
          </pkg:xmlData>
        </pkg:part>
        <xsl:if test="contains(@Target, 'sheet')">
           <xsl:variable name="rel" select="'xl/worksheets/_rels/' || substring($path, 15) || '.rels'"/>
          <pkg:Part pkg:name="/{$rel}">
             <pkg:xmlData>
                <xsl:sequence select="parse-xml(archive:extract-text($zip, $rel))" use-when="function-available('archive:extract-text')"/>
                <xsl:sequence select="zip:xml-entry($zip, $rel)" use-when="function-available('zip:xml-entry')"/>
             </pkg:xmlData>
          </pkg:Part>
        </xsl:if>
      </xsl:for-each>
    </pkg:package>
  </xsl:template>
</xsl:stylesheet>