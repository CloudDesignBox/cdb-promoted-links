<h1>Cloud Design Box Promoted Links Web Part</h1>
<p>This web part replicates the classic Promoted Links Web Part but with added features such as web part properties to change the background colour, size of background image and to select which promoted link list to use. </p>
<p>For more information onthe full Cloud Design Box learning platform for modern SharePoint or custom workflows and design, <a href="https://www.clouddesignbox.co.uk">Contact us via the website.</a> </p>
<img src="https://github.com/CloudDesignBox/cdb-promoted-links/blob/master/preview.gif" alt="preview" />

<h1>Download and install latest version</h1>

<p><a href="https://github.com/CloudDesignBox/cdb-promoted-links/raw/master/Installation.zip">Click here to download install files</a></p>

<b>Installation instructions</b>
<p>- Unzip the file and open the installation folder</p>
<p>-	Upload folder “CDBPromotedLinks” (including the folder itself) into the SiteAssets library at the root site collection (e.g. contoso.sharepoint.com/siteassets).</p>
<p>-	Upload the “cloud-design-box-promoted-links.sppkg” file to the AppCatalog, tick the box to make available on all site collections.
</p>
<b>Updating instructions</b>
<p>Follow the installation instructions above but replace all the files and check in the sppkg file after upload. Please note that upgrading to version 3 will require you to remove and add the web parts onto the page (due to the large configuration and rewrite).</p>

<h1>Working with the code</h1>
### Building the code

```bash
git clone the repo
npm install
# Ensure promoted links list exists on test site
gulp serve
```
