<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:xs="http://www.w3.org/2001/XMLSchema"
	xmlns:sheet="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	xmlns="http://www.tei-c.org/ns/1.0"
	exclude-result-prefixes="xs"
	version="2.0">
	
	<xsl:output indent="yes" />
	<xsl:param name="heading" select="false()" />
	<xsl:param name="title" />
	<xsl:param name="filename" />
	
	<xsl:template match="/">
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
					<xsl:apply-templates select="//sheet:worksheet/sheet:sheetData" />
				</body>
			</text>
		</TEI>
	</xsl:template>
	
	<xsl:template match="sheet:sheetData">
		<table cols="{count(preceding-sibling::sheet:cols/sheet:col)}" rows="{count(sheet:row[descendant::sheet:v])}">
			<xsl:apply-templates select="sheet:row[descendant::sheet:v]" />
		</table>
	</xsl:template>
	
	<xsl:template match="sheet:row">
		<row>
			<xsl:variable name="free"
					select="ancestor::sheet:sheetData/preceding-sibling::sheet:sheetViews/sheet:sheetView/sheet:pane[@state='frozen']/@topLeftCell" />
			<xsl:if test="(string-length($free) &gt; 0 and following-sibling::sheet:row/sheet:c[@r=$free])
				or (not(preceding-sibling::sheet:row) and $heading)">
				<xsl:attribute name="role">label</xsl:attribute>
			</xsl:if>
			<xsl:apply-templates select="sheet:c" />
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