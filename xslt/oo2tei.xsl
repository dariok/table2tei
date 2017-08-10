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
						<p>Created from <xsl:value-of select="$filename" /> by table2TEI, oo2TEI.xsl on <xsl:value-of select="current-dateTime()"/></p>
					</sourceDesc>
				</fileDesc>
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
			<head><xsl:apply-templates select="@table:name"/></head>
			<xsl:apply-templates select="table:table-row" />
		</table>
	</xsl:template>
	
	<xsl:template match="table:table-row[not(table:table-cell)]" />
	<xsl:template match="table:table-row[table:table-cell]">
		<row>
			<xsl:if test="$heading and position() = 1">
				<xsl:attribute name="role">label</xsl:attribute>
			</xsl:if>
			<xsl:apply-templates select="table:table-cell" />
		</row>
	</xsl:template>
	
	<!-- do not handle empty cells at the end -->
	<xsl:template match="table:table-cell[not(text:p or following-sibling::table:table-cell)]"/>
	<xsl:template match="table:table-cell[text:p or following-sibling::table:table-cell]">
		<xsl:variable name="id" select="generate-id()" />
		<cell>
			<xsl:if test="@table:number-columns-spanned">
				<xsl:attribute name="xml:id" select="$id" />
			</xsl:if>
			<xsl:apply-templates select="text:p" />
		</cell>
		<xsl:if test="@table:number-columns-repeated and following-sibling::table:table-cell">
			<xsl:variable name="content">
				<xsl:apply-templates select="text:p" />
			</xsl:variable>
			<xsl:for-each select="1 to xs:integer(@table:number-columns-repeated) - 1">
				<cell>
					<!-- in case of columnspan, give all cells the same content as I have -->
					<xsl:value-of select="$content"/>
				</cell>
			</xsl:for-each>
		</xsl:if>
		<xsl:if test="@table:number-columns-spanned">
			<xsl:variable name="content">
				<xsl:apply-templates select="text:p" />
			</xsl:variable>
			<xsl:for-each select="1 to xs:integer(@table:number-columns-spanned) - 1">
				<cell sameAs="{$id}" />
			</xsl:for-each>
		</xsl:if>
	</xsl:template>
</xsl:stylesheet>