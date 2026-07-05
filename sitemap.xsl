<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:sm="http://www.sitemaps.org/schemas/sitemap/0.9">
<xsl:output method="html" encoding="UTF-8" indent="yes"/>
<xsl:template match="/">
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <title>Sitemap — WordWise PH</title>
  <link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;700;800;900&amp;display=swap" rel="stylesheet"/>
  <style>
    *{margin:0;padding:0;box-sizing:border-box}
    :root{--navy:#1D2B55;--pink:#D4537E;--gold:#FAC775;--light:#F4F6FF}
    body{font-family:Nunito,sans-serif;background:var(--light);color:var(--navy);min-height:100vh}
    header{background:linear-gradient(135deg,#0F1A3E,#1D2B55,#2B1F5E);padding:40px 24px 32px;text-align:center}
    .logo{font-size:28px;font-weight:900;color:#fff;margin-bottom:6px}
    .logo span{color:var(--gold)}
    header p{font-size:13px;color:rgba(255,255,255,0.6);font-weight:700}
    .badge{display:inline-block;background:rgba(250,199,117,0.15);border:1.5px solid rgba(250,199,117,0.35);color:var(--gold);font-size:11px;font-weight:900;padding:4px 14px;border-radius:99px;margin-top:10px}
    main{max-width:700px;margin:32px auto;padding:0 16px 48px}
    .card{background:#fff;border-radius:16px;border:1.5px solid #EAEDF5;overflow:hidden;box-shadow:0 4px 24px rgba(29,43,85,0.07)}
    .card-header{background:var(--navy);padding:14px 20px;display:flex;align-items:center;justify-content:space-between}
    .card-header h2{color:#fff;font-size:14px;font-weight:900}
    .card-header span{background:var(--pink);color:#fff;font-size:11px;font-weight:900;padding:3px 12px;border-radius:99px}
    .url-row{display:flex;align-items:center;padding:13px 20px;border-bottom:1px solid #F1F3F9;transition:background 0.15s}
    .url-row:last-child{border-bottom:none}
    .url-row:hover{background:#F9FAFB}
    .url-icon{width:32px;height:32px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:15px;flex-shrink:0;margin-right:12px}
    .url-link{flex:1;font-size:13px;font-weight:800;color:var(--navy);text-decoration:none;word-break:break-all}
    .url-link:hover{color:var(--pink)}
    .url-freq{font-size:10px;font-weight:900;color:#9CA3AF;background:#F1F3F9;padding:2px 8px;border-radius:99px;flex-shrink:0;margin-left:8px}
    .visit-btn{display:inline-block;font-size:11px;font-weight:900;color:var(--pink);background:#FFF0F7;padding:4px 12px;border-radius:99px;text-decoration:none;margin-left:8px;flex-shrink:0}
    .visit-btn:hover{background:var(--pink);color:#fff}
    footer{text-align:center;font-size:12px;color:#9CA3AF;padding:0 16px 32px}
    footer a{color:var(--pink);font-weight:800;text-decoration:none}
  </style>
</head>
<body>
  <header>
    <div class="logo">WordWise<span>PH</span></div>
    <p>English learning games for Filipino kids</p>
    <div class="badge">&#x1F5FA;&#xFE0F; Site Map</div>
  </header>
  <main>
    <div class="card">
      <div class="card-header">
        <h2>All Pages</h2>
        <span><xsl:value-of select="count(sm:urlset/sm:url)"/> URLs</span>
      </div>
      <xsl:for-each select="sm:urlset/sm:url">
        <xsl:variable name="url" select="sm:loc"/>
        <xsl:variable name="path" select="substring-after($url, 'wordwiseph.com/')"/>
        <xsl:variable name="icon">
          <xsl:choose>
            <xsl:when test="$path = '' or $path = 'index.html'">&#x1F3E0;</xsl:when>
            <xsl:when test="$path = 'app.html'">&#x1F3AE;</xsl:when>
            <xsl:when test="$path = 'about.html'">&#x2139;&#xFE0F;</xsl:when>
            <xsl:when test="$path = 'blog.html'">&#x1F4DA;</xsl:when>
            <xsl:when test="starts-with($path, 'blog-')">&#x1F4DD;</xsl:when>
            <xsl:when test="starts-with($path, 'game-')">&#x1F3B2;</xsl:when>
            <xsl:when test="$path = 'printables.html'">&#x1F5A8;&#xFE0F;</xsl:when>
            <xsl:when test="$path = 'privacy.html'">&#x1F512;</xsl:when>
            <xsl:when test="$path = 'terms.html'">&#x1F4DC;</xsl:when>
            <xsl:otherwise>&#x1F4C4;</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="bg">
          <xsl:choose>
            <xsl:when test="$path = '' or $path = 'index.html'">#FFF0F7</xsl:when>
            <xsl:when test="$path = 'app.html'">#F0F7FF</xsl:when>
            <xsl:when test="starts-with($path, 'blog')">#FFF8E7</xsl:when>
            <xsl:when test="starts-with($path, 'game-')">#F0FFF4</xsl:when>
            <xsl:when test="$path = 'printables.html'">#F5F0FF</xsl:when>
            <xsl:otherwise>#F4F6FF</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <div class="url-row">
          <div class="url-icon" style="background:{$bg}"><xsl:value-of select="$icon" disable-output-escaping="yes"/></div>
          <a class="url-link" href="{$url}"><xsl:value-of select="$url"/></a>
          <xsl:if test="sm:changefreq">
            <span class="url-freq"><xsl:value-of select="sm:changefreq"/></span>
          </xsl:if>
          <a class="visit-btn" href="{$url}" target="_blank">Visit &#x2192;</a>
        </div>
      </xsl:for-each>
    </div>
  </main>
  <footer>
    <p>&#169; 2026 <a href="https://wordwiseph.com">WordWise PH</a> &middot; Sitemap for search engines and humans</p>
  </footer>
</body>
</html>
</xsl:template>
</xsl:stylesheet>