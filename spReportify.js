let spReportifyData = {
    config: {},         // Configurations
    stack: [],          // Stack of functions to run.
    params: {},         // Parameters from the URL. 
    sp: {
        ctx: {},        // SharePoint Context.
        web: {},        // SharePoint get_web().
        caml: {}        // SharePoint CAML Query.
    },
    user: {},           // Name and Email of Runtime user. 
    site: {},           // Info about the SharePoint Site.
    builder: {}         // Work Data for the Report Builder.
}
let _api;

const spReportify = {

/**
 * init
 * Initialize the spReportify Environment.
 */
init: function(InitMode = "builder"){
    this.logBigTitle();
    this.logTitle( `Initializing spReportity in "${InitMode}" mode...` );
    // Load Configurations
        spReportifyData.config = {
            url: _URL,
            reportListName: _REPORT_LIST,
            allowLibraries: _ALLOW_LIBRARY,
            allowLists: _ALLOW_LIST,
            allowHiddenLibraries: _ALLOW_LIBRARY_HIDDEN,
            allowHiddenLists: _ALLOW_LIST_HIDDEN,
            alwaysAllowFields: _ALLOW_FIELDS_ALWAYS
        }
    // SharePoint Site Context
        spReportifyData.sp.ctx = new SP.ClientContext( spReportifyData.config.url );
        spReportifyData.sp.web = spReportifyData.sp.ctx.get_web();
        spReportifyData.sp.ctx.load( spReportifyData.sp.web );

    // Stack Commands
        this.waitingCreate();
        this.waitingShow( "Starting..." );
        this.stackAdd( this.getUser );
        this.stackAdd( this.getSite );

        // Commands for Report Builder
        if( InitMode == "builder" ){
            spReportifyData.sp.caml = new SP.CamlQuery();
            this.stackAdd( this.getLists );
            this.stackAdd( this.initBuilder );
            this.stackAdd( this.builderToggleMode );
        }
        
        // Commends for Report Runner
        if( InitMode == "runner" ){
            //...
        }

    // Start Command Stack
        this.stackRun();
},

/**
 * getUser
 * Retrieve Current User basic Information.
 */
getUser: function(){
    _api = spReportifyData.sp.web.get_currentUser();
    spReportifyData.sp.ctx.load( _api );
    spReportifyData.sp.ctx.executeQueryAsync(
        // Success
        function(){
        spReportifyData.user = {
            email: _api.get_email(),
            title: _api.get_title()
        };
        spr.logTitle( "Runtime User Information" );
        console.table( spReportifyData.user );
        spr.stackRun();
        },
        // Fail
        function(sender, args){
            console.error('Unable to getUser() - ' + args.get_message() );
        }
    );
},

/**
 * getSite
 * Retrieve SP Site Information
 */
getSite: function(){
    spReportifyData.site = {
        title: spReportifyData.sp.web.get_title()
    };
    spr.logTitle( "SharePoint Site Information" );
    console.table( spReportifyData.site );
    spr.stackRun();
},

/**
 * getLists
 * Retrieve all allowed lists & libraries in the site.
 */
getLists: function(){
    spReportifyData.builder["lists"] = [];
    _api = spReportifyData.sp.web.get_lists();
    spReportifyData.sp.ctx.load( _api );
    spReportifyData.sp.ctx.executeQueryAsync(
        // Success
        function(){
            var enumLists = _api.getEnumerator();
            while( enumLists.moveNext() ){
                var thisList = enumLists.get_current();
                // Validate if we should keep this list, using values from Configuration.
                var thisListHidden = thisList.get_hidden();
                var thisListType = ( thisList.get_baseType() == 0 ? "List":"Library");
                var thisListName = thisList.get_title();
                var thisListId = thisList.get_id().toString();
                
                var KeepThisList = true;

                if( thisListType == "Library" ){
                    if( spReportifyData.config.allowLibraries == false ){
                        KeepThisList = false;
                    }else{
                        if( spReportifyData.config.allowHiddenLibraries == false && thisListHidden == true ){
                            KeepThisList = false;
                        }
                    }
                }

                if( thisListType == "List" ){
                    if( spReportifyData.config.allowLists == false ){
                        KeepThisList = false;
                    }else{
                        if( spReportifyData.config.allowHiddenLists == false && thisListHidden == true ){
                            KeepThisList = false;
                        }
                    }
                }

                // Add the list to the spReportifyData.builder.lists Array
                if( KeepThisList ){
                    spReportifyData.builder.lists.push(
                        {
                            id: thisListId,
                            name: thisListName,
                            type: thisListType,
                            hidden: thisListHidden
                        }
                    )
                }
            }
            spr.logTitle( "Available Lists" );
            console.table( spReportifyData.builder.lists );
            spr.stackRun();
        },
        // Failure
        function( sender, args ){
            console.error('Unable to getLists() - ' + args.get_message() );
        }
    )
},



/**
 * initBuilder
 * Prepare the Report Builder.
 */
    initBuilder: function(){
        // Add Options in the List Selector.
            var ctlListSelect = document.getElementById("lstDatasource");
            ctlListSelect.options.length = 0;
            for( const list of spReportifyData.builder.lists ){
                var ctlListOption = document.createElement('option');
                ctlListOption.value = list.id;
                ctlListOption.innerHTML = list.name;
                ctlListSelect.appendChild( ctlListOption );
            }
        // Add a "Select List" Option
            var ctlListOption = document.createElement('option');
            ctlListOption.value = null;
            ctlListOption.innerHTML = "Select a datasource...";
            ctlListSelect.prepend( ctlListOption );
            ctlListSelect.value = null;

            document.getElementById("lstDatasource").addEventListener("change", spReportify.builderUpdateList );
        // Initial Builder Mode - Setting "add" and calling builderToggleMode to switch to "edit".
            spReportifyData.builder["mode"] = "add";
        
        
            //...

        spr.stackRun();
    },


    /**
     * builderFeedReportlist
     * Add Reports in the Report List.
     */
    builderFeedReportList: function(){
        var ctlReportSelect = document.getElementById("lstReport");
        ctlReportSelect.options.length = 0;
        for( const report of spReportifyData.builder.reports ){
            var ctlReportOption = document.createElement('option');
            ctlReportOption.value = report.id;
            ctlReportOption.innerHTML = report.title;
            ctlReportSelect.appendChild( ctlReportOption );
        }
        // Add a "Select Report" Option
        var ctlReportOption = document.createElement('option');
        ctlReportOption.value = null;
        ctlReportOption.innerHTML = "Select a report...";
        ctlReportSelect.prepend( ctlReportOption );
        ctlReportSelect.value = null;

        document.getElementById("lstReport").addEventListener("change", spReportify.builderLoadReport );
        document.getElementById("SelectReport").style.setProperty("display", "block");
    
        //...

    spr.stackRun();
    },

/**
 * builderUpdateList
 * Sequence to run after user changed the List.
 */
builderUpdateList: function(){
    spr.waitingShow( "Fetching Columns..." );
    spr.stackAdd( spr.builderGetColumns );
    spr.stackAdd( spr.builderGetReports );
    spr.stackAdd( spr.builderFeedReportList );
    spr.stackRun();
},

/**
 * builderGetColumns
 * Retrieve Columns for the selected list.
 */
builderGetColumns: function(){
    var listId = document.getElementById("lstDatasource").value
    if( listId != "undefined" && listId+"" != "" ){
        console.log( "Updating Builder for List ID '" + listId + "'" )
        spReportifyData.builder["list"] = {};
        spReportifyData.builder["list"] = spReportifyData.builder.lists.filter(x => x.id == listId)[0];
        spReportifyData.builder.list["id"] = listId;

        spr.logTitle( "Selected List" );
        console.table( spReportifyData.builder.list );

        _api = spReportifyData.sp.web.get_lists().getById(listId).get_fields();
        spReportifyData.sp.ctx.load( _api );
        spReportifyData.sp.ctx.executeQueryAsync(
            // Success
            function(){
                spReportifyData.builder["columns"] = [];
                var enumFields = _api.getEnumerator();
                while( enumFields.moveNext() ){
                    var thisField = enumFields.get_current();

                    var thisFieldTitle = thisField.get_title();
                    var thisFieldStatic = thisField.get_staticName();
                    var thisFieldSealed = thisField.get_sealed();
                    var thisFieldHidden = thisField.get_hidden();
                    var thisFieldFromBase = thisField.get_fromBaseType();

                    var KeepThisField = true;

                    if( thisFieldSealed == true || thisFieldHidden == true ){
                        KeepThisField = false;
                    }
                    if( thisFieldFromBase == true && thisFieldStatic != "ID" ){
                        KeepThisField = false;
                    }



                    // Override the KeepThisField if the Name is in the _ALLOW_FIELDS_ALWAYS config.
                    if( thisFieldStatic in spReportifyData.config.alwaysAllowFields ){
                        KeepThisField = true;
                    }

                    // Add the Column (field) to spReportifyData.builder.columns Array 
                    if( KeepThisField == true ){
                        spReportifyData.builder.columns.push({
                            name: thisFieldStatic,
                            title: thisFieldTitle,
                            type: thisField.get_typeAsString(),
                            description: thisField.get_description()
                        });
                    }

                }
                // Sort the list of columns
                    spReportifyData.builder.columns.sort((a, b) => {
                        const nameA = a.title.toLowerCase();
                        const nameB = b.title.toLowerCase();
                        if( nameA < nameB ){ return -1; }
                        if (nameA > nameB ){ return 1;  }
                        return 0;
                    });
                
                spr.logTitle( "Available Columns in the Selected List" );
                console.table( spReportifyData.builder.columns );
                spr.stackRun();
            },
            // Failure
            function( sender, args ){
                console.error('Unable to get Columns from builderUpdateList() - ' + args.get_message() );
            }
        );
        
    }
},

/**
 * builderGetReports
 * Retrieve Reports for the selected list.
 */
builderGetReports: function(){
    var listId = document.getElementById("lstDatasource").value
    if( listId != "undefined" && listId+"" != "" ){
        spReportifyData.sp.caml.set_viewXml('<View><Query><Where><Eq><FieldRef Name=\'ListId\'/><Value Type=\'Text\'>' + listId + '</Value></Eq></Where></Query></View>');
        _api = spReportifyData.sp.web.get_lists().getByTitle(spReportifyData.config.reportListName).getItems( spReportifyData.sp.caml );
        spReportifyData.sp.ctx.load( _api );
        spReportifyData.sp.ctx.executeQueryAsync(
            // Success
            function(){
                spReportifyData.builder["reports"] = [];
                var enumReports = _api.getEnumerator();
                while( enumReports.moveNext() ){
                    var thisReport = enumReports.get_current();
                    var thisReportId = thisReport.get_id();
                    var thisReportTitle = thisReport.get_item("Title");
                    var thisReportDescription = thisReport.get_item("Description");
                    var thisReportSelect = thisReport.get_item("SelectEntries");
                    var thisReportSort = thisReport.get_item("SortEntries");
                    var thisReportShow = thisReport.get_item("ShowEntries");
                    var thisReportQuery = thisReport.get_item("Query")


                    spReportifyData.builder.reports.push({
                        id: thisReportId,
                        title: thisReportTitle,
                        description: thisReportDescription,
                        select: thisReportSelect,
                        sort: thisReportSort,
                        show: thisReportShow,
                        query: thisReportQuery
                    });
                };

                // Sort the list of reports
                    spReportifyData.builder.reports.sort((a, b) => {
                        const nameA = a.title.toLowerCase();
                        const nameB = b.title.toLowerCase();
                        if( nameA < nameB ){ return -1; }
                        if (nameA > nameB ){ return 1;  }
                        return 0;
                    });

                spr.logTitle( "Available Reports in the Selected List" );
                console.table( spReportifyData.builder.reports );
                spr.stackRun();
            },
            // Failure
            function( sender, args ){
                console.error('Unable to get Reports from builderGetReports() - ' + args.get_message() );
            }
        );
        
    }
},

/**
 * builderToggleMode
 * Switch between Edit an existing report and Create a new report.
 */
builderToggleMode: function(){
    spReportifyData.builder.mode = ( spReportifyData.builder.mode == "edit" ? "add" : "edit" );
    spr.logTitle( "Changing Builder Mode" );
    console.log( `Current Mode: "${spReportifyData.builder.mode}"` );
    if( spReportifyData.builder.mode == "edit" ){
        document.getElementById("tabSwitchMode").innerHTML = "Create a new report...";
        document.getElementById("SelectReportExisting").style.setProperty("display", "block");
        document.getElementById("SelectReportCreate").style.setProperty("display", "none");
    }else{
        document.getElementById("tabSwitchMode").innerHTML = "Edit an existing report...";
        document.getElementById("SelectReportExisting").style.setProperty("display", "none");
        document.getElementById("SelectReportCreate").style.setProperty("display", "block");
    }
    console.log( `New Mode: "${spReportifyData.builder.mode}"` );
    spr.stackRun();
},

/**
 * Load a Report in the Builder.
 */
builderLoadReport: function(){
    spReportifyData.builder.report = {};
    if( spReportifyData.builder.mode == "edit" ){
        spReportifyData.builder.report  = spReportifyData.builder.reports.filter(x => x.id == document.getElementById("lstReport").value )[0];
    }else{
        spReportifyData.builder.report = {
            id: -1,
            title: (document.getElementById("txtNewReportName").value).trim(),
            select: null,
            sort: null,
            show: null,
            description: null,
            query: null
        }
    }

    spr.logTitle(`Report to ${spReportifyData.builder.mode}`);
    console.table( spReportifyData.builder.report );

    // Draw the Builder Interface to display Options
        // Hide the Report Selection Form


        // Show the Editor Interface
            document.getElementById("ReportBuilderSection").style.setProperty("display", "block");



},


/**
 * stackAdd
 * Add a function to the 'stack'.
 */
    stackAdd: function(fctName, fctArgs = null){
        spReportifyData.stack.push({fct: fctName, args: fctArgs});
    },
/**
 * stackRun
 * Execute the next function in the 'stack'.
 */
    stackRun: function(){
        if( spReportifyData.stack.length == 0 ){
            spr.logMute("Command stack is empty!");
            spr.waitingHide();
            return;
        };
        var thisStackEntry = spReportifyData.stack.shift();
        var thisFunction = thisStackEntry.fct;
        var thisArgs = thisStackEntry.args;
        if( thisArgs == null ){
            thisFunction();
        }else{
            thisFunction(thisArgs);
        }
    },

/**
 * getParams
 * Get Parameters from the URL and store them in this.params
 */
    getParams: function(){
        const urlParams = new URLSearchParams(window.location.search);
        const entries = urlParams.entries();
        for( const entry of entries ){
            spReportifyData.params[entry[0].toLowerCase()] = entry[1];
        }
        this.stackRun();
    },


/**
 * waitingCreate
 * Add a waiting DIV in the Page to display when the script is working.
 */
    waitingCreate: function(){
        var ctlDivWaiting = document.createElement('div');
        ctlDivWaiting.id = "wip";
        ctlDivWaiting.style.setProperty("position", "fixed");
        ctlDivWaiting.style.setProperty("width", "100%");
        ctlDivWaiting.style.setProperty("top", "15%");
        ctlDivWaiting.style.setProperty("height", "350px");
        ctlDivWaiting.style.setProperty("left", "0");
        ctlDivWaiting.style.setProperty("justify-content", "center");
        ctlDivWaiting.style.setProperty("vertical-align", "middle");
        ctlDivWaiting.style.setProperty("display", "flex");
        ctlDivWaiting.style.setProperty("background-color", "yellow");
        ctlDivWaiting.innerHTML = `<span id="wipTell">One moment please...</span>`;
        document.body.appendChild( ctlDivWaiting );
        this.stackRun();
    },
/**
 * waitingShow
 * Display a message in the "Please Wait..." popup.
 */
    waitingShow: function( Message = "" ){
        document.getElementById( "wipTell" ).innerHTML = (Message == "" ? "One moment please..." : Message );
        document.getElementById( "wip" ).style.setProperty("display", "block");
    },
/**
 * waitingHide
 * Hide the "Please Wait..." popup.
 */
    waitingHide: function(){
        document.getElementById( "wip" ).style.setProperty("display", "none");
    },

/**
 * logBigTitle
 * Display a Thank you / Welcome message in the Console, for branding purposes.
 */
    logBigTitle: function(){
        console.log(`%cThank you for using spReportify!`, "color: #0C4767;font-family:Tahoma;font-weight:bold;font-size:24px;");
    },
/**
 * logTitle
 * Display a title in the Console.
 */
    logTitle: function( Title ){
        console.log(`%c${Title}`, "color: #233E82;font-family:Tahoma;font-weight:bold;font-size:18px;");
    },
/**
 * logMute
 * Display an entry in the Console, light gray.
 */
    logMute: function( Message ){
    console.log(`%c  ${Message}  `, "background: #D6E5E3;color:#04080F;font-family:Arial;");
}



    


};

let spr = spReportify;

$(document).ready(function(){
    spReportify.init();
 });