<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:archive="http://expath.org/ns/archive"
  xmlns:file="http://expath.org/ns/file"
  xmlns:zip="http://expath.org/ns/zip"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:sheet="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:err="http://www.w3.org/2005/xqt-errors"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  exclude-result-prefixes="#all"
  version="3.0">
  
  <xsl:output indent="1" />
  <xsl:param name="filename" />
  
  <xsl:template name="extract">
    <pkg:package
      xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <xsl:variable name="sheets">
        <xsl:choose>
          <xsl:when test="function-available('archive:extract-text')">
            <xsl:variable name="file" select="file:read-binary($filename)" />
            <xsl:sequence select="parse-xml(archive:extract-text($file, 'xl/workbook.xml'))" use-when="function-available('archive:extract-text')" />
          </xsl:when>
          <xsl:when test="function-available('zip:xml-entry')">
            <xsl:variable name="sheets" select="zip:xml-entry($filename, 'xl/workbook.xml')" use-when="function-available('zip:xml-entry')"/>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>
      <pkg:part pkg:name="xl/workbook.xml">
        <xsl:sequence select="$sheets" />
      </pkg:part>
      <pkg:part pkg:name="xl/_rels/workbook.xml.rels">
        <xsl:choose>
          <xsl:when test="function-available('archive:extract-text')">
            <xsl:variable name="file" select="file:read-binary($filename)" />
            <xsl:sequence select="parse-xml(archive:extract-text($file, 'xl/_rels/workbook.xml.rels'))" use-when="function-available('archive:extract-text')" />
          </xsl:when>
          <xsl:when test="function-available('zip:xml-entry')">
            <xsl:variable name="sheets" select="zip:xml-entry($filename, 'xl/_rels/workbook.xml.rels')" use-when="function-available('zip:xml-entry')"/>
          </xsl:when>
        </xsl:choose>
      </pkg:part>
      <pkg:part pkg:name="xl/sharedStrings.xml">
        <xsl:choose>
          <xsl:when test="function-available('archive:extract-text')">
            <xsl:variable name="file" select="file:read-binary($filename)" />
            <xsl:sequence select="parse-xml(archive:extract-text($file, 'xl/sharedStrings.xml'))" use-when="function-available('archive:extract-text')" />
          </xsl:when>
          <xsl:when test="function-available('zip:xml-entry')">
            <xsl:variable name="sheets" select="zip:xml-entry($filename, 'xl/sharedStrings.xml')" use-when="function-available('zip:xml-entry')"/>
          </xsl:when>
        </xsl:choose>
      </pkg:part>
      <xsl:apply-templates select="$sheets//sheet:sheets/sheet:sheet" />
    </pkg:package>
  </xsl:template>
  
  <xsl:template match="sheet:sheet">
    <xsl:variable name="name">
      <xsl:value-of select="'xl/worksheets/sheet' || @sheetId || '.xml'" />
    </xsl:variable>
    <xsl:variable name="sheet">
      <xsl:choose>
        <xsl:when test="function-available('archive:extract-text')">
          <xsl:variable name="file" select="file:read-binary($filename)" />
          <xsl:sequence select="parse-xml(archive:extract-text($file, $name))" use-when="function-available('archive:extract-text')" />
        </xsl:when>
        <xsl:when test="function-available('zip:xml-entry')">
          <xsl:variable name="sheets" select="zip:xml-entry($filename, $name)" use-when="function-available('zip:xml-entry')"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>
    <pkg:part pkg:name="{$name}"
      xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <xsl:sequence select="$sheet" />
    </pkg:part>
    
    <xsl:variable name="nameRel">
      <xsl:value-of select="'xl/worksheets/_rels/sheet' || @sheetId || '.xml.rels'" />
    </xsl:variable>
    <xsl:variable name="sheet">
      <xsl:choose>
        <xsl:when test="function-available('archive:extract-text')">
          <xsl:variable name="file" select="file:read-binary($filename)" />
          <xsl:sequence select="parse-xml(archive:extract-text($file, $nameRel))" use-when="function-available('archive:extract-text')" />
        </xsl:when>
        <xsl:when test="function-available('zip:xml-entry')">
          <xsl:variable name="sheets" select="zip:xml-entry($filename, $nameRel)" use-when="function-available('zip:xml-entry')"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>
    <pkg:part pkg:name="{$nameRel}"
      xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <xsl:sequence select="$sheet" />
    </pkg:part>
  </xsl:template>
</xsl:stylesheet>