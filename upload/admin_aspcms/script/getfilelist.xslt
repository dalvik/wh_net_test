<?xml version="1.0" encoding="gbk"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output omit-xml-declaration="yes"/>
	<!--
	Copyright (C) 2007,2008 NewAsp.Net. All rights reserved.
	Written by newasp.net Sunwin
	-->
<xsl:template match="/">
	<xml>
	<xsl:for-each select="xml/row">
		<xsl:sort select="@datelastmodified" order="descending"/>
		<xsl:element name="row">
			<xsl:attribute name="position"><xsl:value-of select="position()"/></xsl:attribute>
			<xsl:attribute name="name"><xsl:value-of select="@name"/></xsl:attribute>
			<xsl:attribute name="type"><xsl:value-of select="@type"/></xsl:attribute>
			<xsl:attribute name="size"><xsl:value-of select="@size"/></xsl:attribute>
			<xsl:attribute name="datelastmodified"><xsl:value-of select="@datelastmodified"/></xsl:attribute>
			<xsl:attribute name="datecreated"><xsl:value-of select="@datecreated"/></xsl:attribute>
		</xsl:element>
	</xsl:for-each>
	</xml>
</xsl:template>
</xsl:stylesheet>