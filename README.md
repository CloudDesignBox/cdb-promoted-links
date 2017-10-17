## cloud-design-box-promoted-links

<h1>Download and install version 2</h1>

<p><a href="https://github.com/CloudDesignBox/cdb-promoted-links/blob/master/InstallationV1.zip">Click here to download install files</a></p>

<b>Installation instructions</b>
- Unzip the file and open the V2 installation folder
-	Upload folder “CDBPromotedLinks” (including the folder itself) into the SiteAssets library at the root site collection (e.g. contoso.sharepoint.com/siteassets). This is so that it doesn’t use a CDN.
-	Upload the “cloud-design-box-promoted-links.sppkg” file to the AppCatalog, tick the box to make available on all site collections.
This is where you include your WebPart documentation.

### Building the code

```bash
git clone the repo
npm install
# Ensure promoted links list exists on test site
gulp serve
```
