<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="2.0"
                xmlns:kiln="http://www.kcl.ac.uk/artshums/depts/ddh/kiln/ns/1.0"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:i18n="http://apache.org/cocoon/i18n/2.1">

  <xsl:template match="response" mode="text-index">
    <table class="tablesorter">
      <thead>
        <tr>
          <!-- Let us assume that all texts have a filename, ID, and
               title. -->
          <th><i18n:text i18n:key="file-name">Filename</i18n:text></th>
          <th><i18n:text i18n:key="id">ID</i18n:text></th>
          <th><i18n:text i18n:key="title">Title</i18n:text></th>
          <xsl:if test="result/doc/arr[@name='author']/str">
            <th><i18n:text i18n:key="author">Author</i18n:text></th>
          </xsl:if>
          <xsl:if test="result/doc/arr[@name='editor']/str">
            <th><i18n:text i18n:key="editor">Editor</i18n:text></th>
          </xsl:if>
          <xsl:if test="result/doc/str[@name='publication_date']">
            <th><i18n:text i18n:key="publication-date">Publication Date</i18n:text></th>
          </xsl:if>
        </tr>
      </thead>
      <tbody>
        <xsl:apply-templates mode="text-index" select="result" />
      </tbody>
    </table>
  </xsl:template>

  <xsl:template match="result[not(doc)]" mode="text-index">
    <p>There are no files indexed from
    webapps/ROOT/content/xml/<xsl:value-of select="$document_type" />!
    Put some there, index them from the admin page, and this page will
    become much more interesting.</p>
  </xsl:template>

  <xsl:template match="result/doc" mode="text-index">
    <tr>
      <xsl:apply-templates mode="text-index" select="str[@name='file_path']" />
      <xsl:apply-templates mode="text-index" select="str[@name='document_id']" />
      <xsl:apply-templates mode="text-index" select="arr[@name='document_title']" />
      <xsl:apply-templates mode="text-index" select="arr[@name='author']" />
      <xsl:apply-templates mode="text-index" select="arr[@name='editor']" />
      <xsl:apply-templates mode="text-index" select="str[@name='publication_date']" />
    </tr>
  </xsl:template>

  <xsl:template match="str[@name='file_path']" mode="text-index">
    <xsl:variable name="filename" select="substring-after(., '/')" />
    <td>
      <a href="{kiln:url-for-match($match_id, ($language, $filename), 0)}">
        <xsl:value-of select="$filename" />
      </a>
    </td>
  </xsl:template>

  <xsl:template match="str[@name='document_id']" mode="text-index">
    <td><xsl:value-of select="." /></td>
  </xsl:template>

  <xsl:template match="arr[@name='document_title']" mode="text-index">
    <td><xsl:value-of select="string-join(str, '; ')" /></td>
  </xsl:template>

  <xsl:template match="arr[@name='author']" mode="text-index">
    <td><xsl:value-of select="string-join(str, '; ')" /></td>
  </xsl:template>

  <xsl:template match="arr[@name='editor']" mode="text-index">
    <td><xsl:value-of select="string-join(str, '; ')" /></td>
  </xsl:template>

  <xsl:template match="str[@name='publication_date']">
    <td><xsl:value-of select="." /></td>
  </xsl:template>

</xsl:stylesheet>
