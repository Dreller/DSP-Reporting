/**
 * Configuration File
 */


/**
 * SHAREPOINT SITE URL
 * ------------------------------------------------------------------
 * This is the URL to the SharePoint site where you want to install
 * spReportify.  This value is used to send HTTP requests to get data
 * to report and to generate URLs to Reports.
 * 
 * Controlled transformations:
 *      - with or without "https://" (will be added if missing),
 *      - http will be changed to https,
 *      - code will ensure there is "/" at the end,
 * 
 *      Examples:
 *      - tenant.sharepoint.com/sites/MySite/
 *      - https://tenant.sharepoint.com/sites/MySite
 */
var _SPR_URL = "https://123.sharepoint.com/sites/Mysite/";

/**
 * LIST TO STORE REPORTS
 * ------------------------------------------------------------------
 * This is the name of the SharePoint List where Reports definitions
 * are stored.  There is 2 methods to identify this list, you can
 * use it's GUID or it's name, please, ensure both variables are
 * correctly set.
 * 
 * --- Variable _SPR_REPORTLISTTYPE (ReportListType) ----------------
 * This is a flag to tell if the reference is a GUID or the Title.
 * 
 * Acceptable values:
 *  - "guid"
 *  - "title"
 * 
 * --- Variable _SPR_REPORTLISTREF (ReportListRef) ------------------
 * This is the GUID or the Title of the SharePoint List.
 */
var _SPR_REPORTLISTTYPE = "title";
var _SPR_REPORTLISTREF = "Reports";

/**
 * ALLOW REPORTING ON LIBRARIES
 * ------------------------------------------------------------------
 * Set this value to 'true' if you want to allow using Libraries as
 * datasource for reports.
 */
var _SPR_ALLOWLIBRARY = false;

/**
 * ALLOW REPORTING ON HIDDEN LIBRARIES
 * ------------------------------------------------------------------
 * Set this value to 'true' if you want to allow using hidden
 * libraries as datasource for reports.
 */
var _SPR_ALLOWHIDDENLIBRARY =  false;

/**
 * ALLOW REPORTING ON LISTS
 * ------------------------------------------------------------------
 * Set this to 'true' if you want to allow using Lists as datasources
 * for reports.  THIS SHOULD BE SET TO 'true' FOR NORMAL USE.
 */
var _SPR_ALLOWLIST = true;

/**
 * ALLOW REPORTING ON HIDDEN LISTS
 * ------------------------------------------------------------------
 * Set this to 'true' if you want to allow using hidden Lists as
 * datasources for reports.
 */
var _SPR_ALLOWHIDDENLIST = false;

/**
 * ALWAYS ALLOW FIELDS
 * ------------------------------------------------------------------
 * This is an Array of Fields that should always be available.  Use
 * this parameter to make the field 'ID' always available.  You
 * can also use this parameter for 'Content Type' per example.
 */
var _SPR_ALLOWFIELDS = [
    "ID"
];

/**
 * LOG LEVEL
 * ------------------------------------------------------------------
 * Set the level of details written to the Browser Console.
 * 
 * Accepted values: 
 *  error       Errors only.                                        4
 *  warn        Potential issues that may lead to errors.           3
 *  info        General Information.                                2
 *  trace       Output everything to the console.                   1
 * 
 *  (empty)     Disabled.                                           0
 */
var _SPR_LOGLEVEL = "trace";

/**
 * SCRIPT PAGE NAMES
 * ------------------------------------------------------------------
 * This is the name of the pages for spReportify Script.  This is a
 * more technical parameter, used to generate URLs to share Reports,
 * and to determine if the user is using the Runner or the Builder.
 * 
 * If you change the name of '.aspx' pages, update these parameters.
 * 
 * --- Variable _SPR_PAGEBUILDER (PageBuilder) ----------------------
 * Set this value to the name of the '.aspx' page of the Report
 * Builder.  Example:  'file.aspx'.
 * 
  * --- Variable _SPR_PAGERUNNER (PageRunner) -----------------------
 * Set this value to the name of the '.aspx' page of the Report
 * Runner.  Example:  'file.aspx'.
 * 
 */
var _SPR_PAGEBUILDER = "sprBuild.aspx";
var _SPR_PAGERUNNER = "sprExecute.aspx";