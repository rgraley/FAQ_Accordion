/*
Build a Helloworld webpart: https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part
Connect your webpart to SharePoint: https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/connect-to-sharepoint
Deploy your webpart to a SharePoint Page: https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/serve-your-web-part-in-a-sharepoint-page
Add jQueryUI accordion to your webpart: https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/add-jqueryui-accordion-to-web-part
Webpart Title control for SPFx: https://www.c-sharpcorner.com/article/pnp-webparttitle-control-for-spfx/
SharePoint search using PnP in SPFx Client Webpart: https://social.technet.microsoft.com/wiki/contents/articles/37481.sharepoint-search-using-pnp-in-sharepoint-framework-spfx-client-web-part.aspx

//https://www.techwithnakkeeran.com/2018/01/building-sharepoint-search-queries.html
//https://techcommunity.microsoft.com/t5/sharepoint-developer/pnp-sp-search-gt-add-property-to-searchresults/m-p/115026
Reference: https://social.technet.microsoft.com/wiki/contents/articles/35796.sharepoint-2013-using-rest-api-for-selecting-filtering-sorting-and-pagination-in-sharepoint-list.aspx
Reference: https://www.c-sharpcorner.com/article/spfx-the-modern-way-to-show-loading47progress-indicators-using-react-shimm/

Learning: https://www.3queries.com/tutorials/sharepoint-framework-spfx-tutorial-for-beginners/635/

How to deploy: https://www.c-sharpcorner.com/article/build-and-deploy-the-client-side-web-part-spfx-in-sharepoint-online/


installs used to setup a project
npm install jquery@2
npm install jqueryui
npm install @types/jquery@2 --save-dev
npm install @types/jqueryui --save-dev
npm install @pnp/spfx-controls-react --save
npm install @pnp/pnpjs --save 
npm install sp-pnp-js --save
npm install office-ui-fabric-react --save
*/


//Structure of refinement strings
//refinementfilters=bcTopics:("High Five Rewards*") contains
//refinementfilters=bcTopics:and("Test*", "High Five Rewards*") contains all
//refinementfilters=bcTopics:or("Test*", "High Five Rewards*") may contain any

/*
Read Me:
For this code to work properly there are some environment configurations that need to be made
Site Columns: 
There needs to be a multiline column with the following properties
Internal Name: bcAnswer;
Display Name: Answer
Type: Multiple lines of text
Group: All BayCare Team Member Portal FAQ Columns
Description: Use this question to answer FAQ questions
Required: Yes
Allow unlimited length in documents libraries: No
Number of lines for editing: 6
Specify the type of text to allow: Enhanced rich text(Rich text with pictures, tables, and hyperlinks)
Append Changes to Existing Text: No

Site Content Type:
There needs to be a site content type with the following properties
Name: BayCare FAQs List
Description: 
Parent: Item
Group: All BayCare Team Member Lists
Columns:
    Title: Required
    Answer: Required (see bcAnswer site column created above)
//NOTE: other columns may be added as needed; but these are the required for the solution to work.
//NOTE: Take note of the GUID of this content type for later use

Content Type needs to be added to a List and content can be added as required.

Search Center Administration
Crawled Properties: 
Property Name: ows_bcAnswer
Property Name: ows_r_MTXT_bcAnswer

bcAnswer Mapped Property
Choose an available RefinableString and map it to the crawled property
Example:
    Property name: RefinableString15
    Description: blank
    Alias: bcAnswer
    Add a Mapping: ows_bcAnswer

bcTopics Mapped Property
Choose an available RefinableString and map it to the crawled property
Example:
    Property name: RefinableString16
    Description: blank
    Alias: bcTopics
    Add a Mapping: ows_bcTopics


Manage Result Sources: 
There needs to be a result source with the following properties
Name: BayCare FAQ List Items
Description: This result source retrieves all list items that have a content type of BayCare FAQs List
Protocal: Local SharePoint
Type: SharePoint Search Results
Query Transform: {searchTerms} (ContentTypeId:<USE "BayCare FAQs List" CONTENT TYPE GUID >*)
Credentials Information
    Default Authentication
NOTE: Take note of the sourceid Guid of the result source
This can be found in the edit url of the content type. Example:/_layouts/15/searchadmin/EditResultSource.aspx?level=tenant&sourceid=4225801d-e615-4925-84d8-4f3ec2c22d6a
*/