<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:xs="http://www.w3.org/2001/XMLSchema"
	xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
	xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
	xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
	xmlns="http://www.tei-c.org/ns/1.0"
	exclude-result-prefixes="xs"
	version="2.0">
	
	<xsl:output indent="yes" />
	
	<xsl:template match="/">
		<TEI version="5.0">
			<teiHeader>
				<!-- TODO -->
			</teiHeader>
			<text>
				<body>
					<xsl:apply-templates select="//office:document-content/office:body/office:spreadsheet"/>
				</body>
			</text>
		</TEI>
	</xsl:template>
	
	<xsl:template match="/office:spreadsheet">
		<xsl:apply-templates select="table:table" />
	</xsl:template>
	
	<xsl:template match="table:table">
		<table cols="{count(table:table-column)}" rows="{count(table:table-row)}">
			<xsl:apply-templates select="table:table-row" />
		</table>
	</xsl:template>
	
	<xsl:template match="table:table-row">
		<row>
			<xsl:apply-templates select="table:table-cell" />
		</row>
	</xsl:template>
	
	<xsl:template match="table:table-cell">
		<cell>
			<xsl:apply-templates select="text:p" />
		</cell>
	</xsl:template>
</xsl:stylesheet>