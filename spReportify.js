// Structure of spReportifyData
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
        builder: {},        // Work Data for the Report Builder.
        environ: {},        // Environment (Browser, Language, etc.)
        runner: {}
    }
// Variable to refer to SP API
    let _api;

const spReportify = {
vm: String.fromCharCode(253),
// List of Operators
_op: [
    { op: "eq", label: "Equal", notation: "=", syntax: "seq" },
    { op: "ne", label: "Not equal", notation: "<>", syntax: "seq" },
    { op: "gt", label: "Greater than", notation: ">", syntax: "seq" },
    { op: "ge", label: "Greater than or equal", notation: ">=", syntax: "seq" },
    { op: "lt", label: "Less than", notation: "<", syntax: "seq" },
    { op: "le", label: "Less than or equal", notation: "<=", syntax: "seq" }
],
// List of Directions
_dir: [
        { direction: "asc", label: "Ascending (A-Z)" },
        { direction: "desc", label: "Descending (Z-A)"}
    ],
/**
 * init
 * Initialize the spReportify Environment.
 */
init: function(){
    // Big Title to Console
        this.logBigTitle();

    // Set the Runtime Level
        var LogLevel = 0;
        switch( _SPR_LOGLEVEL ){
            case "error":
                LogLevel = 4;
                break;
            case "warn":
                LogLevel = 3;
                break;
            case "info":
                LogLevel = 2;
                break;
            case "trace":
                LogLevel = 1;
                break;
        }

    // Validate the SharePoint Site URL
        var ThisUrl = new URL( _SPR_URL );
        // Ensure we are using https
            ThisUrl.protocol = "https:";
    
    // Build the Configuration Object
        spReportifyData.config = {
            url: ThisUrl.toString(),
            reportListRefType: _SPR_REPORTLISTTYPE,
            reportListRef: _SPR_REPORTLISTREF,
            allowLibraries: _SPR_ALLOWLIBRARY,
            allowHiddenLibraries: _SPR_ALLOWHIDDENLIBRARY,
            allowLists: _SPR_ALLOWLIST,
            allowHiddenLists: _SPR_ALLOWHIDDENLIST,
            allowFields: _SPR_ALLOWFIELDS,
            logLevel: LogLevel,
            pageBuilder: _SPR_PAGEBUILDER.toLowerCase(),
            pageRunner: _SPR_PAGERUNNER.toLowerCase()
        }
    
    // Load Environment Information
        spReportifyData.environ = {
            language: ((window.navigator.language).split("-")[0]).toLowerCase(),
            timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
            page: ( (window.location.pathname).split("/").pop() ).toLowerCase(),
        };
        // Runtime Mode
            spReportifyData.environ["mode"] = ( spReportifyData.environ.page == spReportifyData.config.pageBuilder ? "builder" : "runner" );

    // Output to Console - Initialization
        spr.logTitle( "Initializing spReportify..." );
        spr.logTrace( `Log Level is: ${(_SPR_LOGLEVEL).toUpperCase()} (${spReportifyData.config.logLevel})` );
        spr.logTrace("Configurations");
        spr.logTable( spReportifyData.config );
        spr.logTrace("Environment");
        spr.logTable( spReportifyData.environ );
    
    // Start the Script
        this.start();
},

start: function(){
    // Log Intro
        spr.logInfo( `Running spReportify in "${spReportifyData.environ.mode}" mode.` );

    // SharePoint Site Context
        spReportifyData.sp.ctx = new SP.ClientContext( spReportifyData.config.url );
        spReportifyData.sp.web = spReportifyData.sp.ctx.get_web();
        spReportifyData.sp.ctx.load( spReportifyData.sp.web );

    // Stack Commands
        this.waitingCreate();
        this.stackAdd( this.getUser );
        this.stackAdd( this.getSite );

        // Commands for Report Builder
            if( spReportifyData.environ.mode == "builder" ){
                this.waitingShow( "Starting Report Builder..." );
                spReportifyData.sp.caml = new SP.CamlQuery();
                this.stackAdd( this.getLists );
                this.stackAdd( this.initBuilder );
                spReportifyData.builder.mode = "edit";
            }
        
        // Commends for Report Runner
            if( spReportifyData.environ.mode == "runner" ){
                this.waitingShow( "Starting Report Runner..." );
                this.stackAdd( this.getParams );
                this.stackAdd( this.runnerGetReport );

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
    spr.waitingShow( "Identifying User..." );
    _api = spReportifyData.sp.web.get_currentUser();
    spReportifyData.sp.ctx.load( _api );
    spReportifyData.sp.ctx.executeQueryAsync(
        // Success
        function(){
        spReportifyData.user = {
            email: _api.get_email(),
            title: _api.get_title()
        };
        spr.logInfo( "Runtime User Information" );
        spr.logTable( spReportifyData.user );
        spr.stackRun();
        }, 
        Function.createDelegate( this, this.logSysError )
    );
},

/**
 * getSite
 * Retrieve SP Site Information
 */
getSite: function(){
    spr.waitingShow( "Getting Source Site Name..." );
    spReportifyData.site = {
        title: spReportifyData.sp.web.get_title()
    };
    spr.logInfo( "SharePoint Site Information" );
    spr.logTable( spReportifyData.site );
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
            spr.logInfo( "Available Lists" );
            spr.logTable( spReportifyData.builder.lists );
            spr.stackRun();
        },
        // Failure
        Function.createDelegate( this, this.logSysError )
    )
},



/**
 * initBuilder
 * Prepare the Report Builder.
 */
    initBuilder: function(){
        // Add Options in the List Selector.
            var ctlListSelect = document.getElementById("BuilderFormControlDatasource");
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

            document.getElementById("BuilderFormControlDatasource").addEventListener("change", spReportify.builderUpdateList );

            // Prepare Modal for the Dictionary
            const modals = document.querySelectorAll("[data-modal]");
            modals.forEach( function( trigger ){
                trigger.addEventListener("click", function( event ){
                    event.preventDefault();
                    const modal = document.getElementById( trigger.dataset.modal );
                    modal.classList.add("open");
                    const exits = modal.querySelectorAll(".modal-exit");
                    exits.forEach( function( exit ) {
                        exit.addEventListener( "click", function( event ){
                            event.preventDefault();
                            modal.classList.remove("open");
                        })
                    });
                }
            )});

            //...

        spr.stackRun();
    },


    /**
     * builderFeedReportlist
     * Add Reports in the Report List.
     */
    builderFeedReportList: function(){
        var ctlReportSelect = document.getElementById("BuilderFormControlReportPicker");
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

       // document.getElementById("BuilderFormControlReportPicker").addEventListener("change", spReportify.builderLoadReport );
        document.getElementById("BuilderFormSectionReportPicker").style.setProperty("display", "block");
        document.getElementById("BuilderFormSectionAction").style.setProperty("display", "block");
        document.getElementById("BuilderFormSectionLoadCreateReport").style.setProperty("display", "block");
        document.getElementById("BuilderFormSectionLoadCreateReport").addEventListener("click", spReportify.builderLoadReport );
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
    var listId = document.getElementById("BuilderFormControlDatasource").value
    if( listId != "undefined" && listId+"" != "" ){
        spReportifyData.builder["list"] = {};
        spReportifyData.builder["list"] = spReportifyData.builder.lists.filter(x => x.id == listId)[0];
        spReportifyData.builder.list["id"] = listId;

        spr.logTrace( "Selected List" );
        spr.logTable( spReportifyData.builder.list );

        _api = spReportifyData.sp.web.get_lists().getById(listId).get_fields();
        spReportifyData.sp.ctx.load( _api );
        spReportifyData.sp.ctx.executeQueryAsync(
            // Success
            function(){
                spReportifyData.builder["columns"] = [];
                var enumFields = _api.getEnumerator();
                while( enumFields.moveNext() ){
                    var thisField = enumFields.get_current();
                    console.log( thisField );
                    var thisFieldTitle = thisField.get_title();
                    var thisFieldStatic = thisField.get_staticName();
                    var thisFieldSealed = thisField.get_sealed();
                    var thisFieldHidden = thisField.get_hidden();
                    var thisFieldFromBase = thisField.get_fromBaseType();
                    var thisFieldSortable = thisField.get_sortable();
                    var thisFieldXML = thisField.get_schemaXml();

                    var KeepThisField = true;

                    if( thisFieldSealed == true || thisFieldHidden == true ){
                        KeepThisField = false;
                    }
                    if( thisFieldFromBase == true && thisFieldStatic != "ID" ){
                        KeepThisField = false;
                    }



                    // Override the KeepThisField if the Name is in the _ALLOW_FIELDS_ALWAYS config.
                    if( thisFieldStatic in spReportifyData.config.allowFields ){
                        KeepThisField = true;
                    }

                    // Add the Column (field) to spReportifyData.builder.columns Array 
                    if( KeepThisField == true ){
                        spReportifyData.builder.columns.push({
                            name: thisFieldStatic,
                            title: thisFieldTitle,
                            type: thisField.get_typeAsString(),
                            description: thisField.get_description(),
                            sortable: thisFieldSortable,
                            indexed: thisField.get_indexed(),
                            schema: thisFieldXML
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
                
                spr.logTrace( "Available Columns in the Selected List" );
                spr.logTable( spReportifyData.builder.columns );

                // Write the Dictionary
                spReportifyData.builder.columns.forEach( function( thisColumn ){
                    var elRow = document.createElement("tr");
                    elRow.innerHTML = `<td>${thisColumn.title}</td><td>${thisColumn.name}</td><td>${thisColumn.type}</td><td>${thisColumn.description}</td>`;
                    document.getElementById("BuilderDictionaryTableBody").appendChild(elRow);
                });
                

                spr.stackRun();
            },
            // Failure
            Function.createDelegate( this, this.logSysError )
        );
        
    }
},

/**
 * builderGetReports
 * Retrieve Reports for the selected list.
 */
builderGetReports: function(){
    var listId = document.getElementById("BuilderFormControlDatasource").value
    if( listId != "undefined" && listId+"" != "" ){
        spReportifyData.sp.caml.set_viewXml('<View><Query><Where><Eq><FieldRef Name=\'ListId\'/><Value Type=\'Text\'>' + listId + '</Value></Eq></Where></Query></View>');
        // Get Reports using Reports List Reference
        switch( (spReportifyData.config.reportListRefType).toLowerCase() ){
            case "title":
                _api = spReportifyData.sp.web.get_lists().getByTitle(spReportifyData.config.reportListRef).getItems( spReportifyData.sp.caml );
                break;
            case "guid":
                _api = spReportifyData.sp.web.get_lists().getById(spReportifyData.config.reportListRef).getItems( spReportifyData.sp.caml );
                break;
        }
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

                spr.logTrace( "Available Reports in the Selected List" );
                spr.logTable( spReportifyData.builder.reports );
                spr.stackRun();
            },
            // Failure
            Function.createDelegate( this, this.logSysError )
        );
        
    }
},

/**
 * builderToggleMode
 * Switch between Edit an existing report and Create a new report.
 */
builderToggleMode: function(TargetMode = 1){
    var NewMode = ( TargetMode == 1 || typeof TargetMode == "undefined" || typeof spReportifyData.builder.mode == "undefined" ? "add" : "edit" );
    if( NewMode == spReportifyData.builder.mode ){ return; }
    spr.logTrace( "Changing Builder Mode" );
    spr.logTrace( `Current Mode: "${spReportifyData.builder.mode}"` );
    spReportifyData.builder.mode = NewMode;
    if( spReportifyData.builder.mode == "edit" ){
        document.getElementById("BuilderFormControlActionChoiceEdit").checked = true;
        document.getElementById("BuilderFormSectionReportPicker").style.setProperty("display", "block");
        document.getElementById("BuilderFormSectionReportNaming").style.setProperty("display", "none");
    }else{
        document.getElementById("BuilderFormControlActionChoiceCreate").checked = true;
        document.getElementById("BuilderFormSectionReportPicker").style.setProperty("display", "none");
        document.getElementById("BuilderFormSectionReportNaming").style.setProperty("display", "block");
    }
    spr.logTrace( `New Mode: "${spReportifyData.builder.mode}"` );
    spr.stackRun();
},

/**
 * Load a Report in the Builder.
 */
builderLoadReport: function(){
    spReportifyData.builder.report = {};
    if( spReportifyData.builder.mode == "edit" ){
        spReportifyData.builder.report = spReportifyData.builder.reports.filter(x => x.id == document.getElementById("BuilderFormControlReportPicker").value )[0];
    }else{
        spReportifyData.builder.report = {
            id: -1,
            title: (document.getElementById("BuilderFormControlReportNaming").value).trim(),
            select: null,
            sort: null,
            show: null,
            description: null,
            query: null
        }
    }

    spr.logTrace(`Report to ${spReportifyData.builder.mode}`);
    spr.logTable( spReportifyData.builder.report );

    // Draw the Builder Interface to display Options
        // Hide the Report Selection Form and show a read-only summary
            spr.hide("BuilderIdentifyReport");
            document.getElementById("BuilderReportIdentityDatasource").innerHTML = spReportifyData.builder.list.name;
            document.getElementById("BuilderReportIdentityReportName").innerHTML = spReportifyData.builder.report.title;
            spr.show("BuilderReportIdentity");

        // Draw Lines in Sections
            if( spReportifyData.builder.report.select != null ){
                spReportifyData.builder.report.select.split("\n").forEach( function( thisRow ){
                    spr.builderDrawRow(1, thisRow );
                })
            }
            if( spReportifyData.builder.report.sort != null ){
                spReportifyData.builder.report.sort.split("\n").forEach( function( thisRow ){
                    spr.builderDrawRow(2, thisRow );
                })
            }
            if( spReportifyData.builder.report.show != null ){
                spReportifyData.builder.report.show.split("\n").forEach( function( thisRow ){
                    spr.builderDrawRow(3, thisRow );
                })
            }

        // Report Options
            document.getElementById("BuilderFormOptionDescription").value = spReportifyData.builder.report.description;

        // Show the Editor Interface
            spr.show("BuilderForm");




},

/**
 * builderValidateReportName
 */
builderValidateReportName: function(){
    var ControlRef = document.getElementById("BuilderFormControlReportNaming");
    var TestedName = ControlRef.value;
    TestedName = TestedName.toLowerCase().replace(/^\s+|\s+$/gm,'');
    if( spReportifyData.builder.reports.filter( x => x.title.toLowerCase() == TestedName).length == 0 ){
        ControlRef.classList.remove("TextBoxError");
        document.getElementById("BuilderFormAlertReportNaming_AlreadyUsed").style.setProperty("display", "none");
        document.getElementById("BuilderFormControlButtonLoadCreate").style.setProperty("display", "block");
    }else{
        ControlRef.classList.add("TextBoxError");
        document.getElementById("BuilderFormAlertReportNaming_AlreadyUsed").style.setProperty("display", "block");
        document.getElementById("BuilderFormControlButtonLoadCreate").style.setProperty("display", "none");
    }
    

},


/**
 * builderDrawRow
 * Add a new Row in the Section.
 * SectionNumber: 1 = Select, 2 = Sort, 3 = Show.
 * RowDefn:  null for a new row.
 */
builderDrawRow: function( SectionNumber, RowDefn = null ){
    spr.logTitle("Add a new Row in Section # " + SectionNumber);
    var RowUID = crypto.randomUUID();
    spr.logTrace('Row UID: ' + RowUID ); 
    // Set the Literal Section Name
    var SectionName = "";
    switch( SectionNumber ){
        case 1:
            SectionName = "Select";
            break;
        case 2:
            SectionName = "Sort";
            break;
        case 3:
            SectionName = "Show";
            break;
    }

    // Parse the RowDefn if not null
    if( RowDefn != null ){
        // Explode the Row
        let RowDef = RowDefn.split( spr.vm );
        switch( SectionNumber ){
            case 1:
            /**
             * Structure of a SELECT Entry.
             * ------------------------------
             * Index    Data
             * -------- ---------------------
             * 0        Kind: (R)egular.
             * 1        Column Static Name.
             * 2        Operator.
             * 3        Value to compare for select.
             */
                RowColumn = RowDef[1];
                RowOperator = RowDef[2];
                RowValue = RowDef[3];
                break;
            case 2:
            /**
             * Structure of a SORT Entry.
             * ------------------------------
             * Index    Data
             * -------- ---------------------
             * 0        Kind: (R)egular.
             * 1        Column Static Name.
             * 2        Direction.
             */
                RowColumn = RowDef[1];
                RowDirection = RowDef[2];
                break;
            case 3:
            /**
             * Structure of a SHOW Entry.
             * ------------------------------
             * Index    Data
             * -------- ---------------------
             * 0        Kind: (R)egular.
             * 1        Column Static Name.
             * 2        Column Header (Label/Title).
             */
                RowColumn = RowDef[1];
                RowLabel = RowDef[2];
                break;
        }
    }else{
        var RowColumn = "";
        var RowOperator = "";
        var RowDirection = "";
        var RowValue = "";
        var RowLabel = ""
    }

    // Create the new row
        var elRow = document.createElement("tr");
        elRow.id = SectionName + "_" + RowUID;

    // Create the Column Selector
        var elSelect = document.createElement("select");
        elSelect.id = "Column_" + RowUID;
        // Insert Columns in the Select
            spReportifyData.builder.columns.forEach( function( thisColumn ){
                // Skip "sortable = false" for SectionNumber = 2
                // Columns needs to be sortable to be used in this section.
                    if( SectionNumber != 2 || thisColumn.sortable == true ){
                        var elOption = document.createElement("option");
                        elOption.value = thisColumn.name;
                        elOption.text = thisColumn.title;
                        elSelect.appendChild( elOption );
                    }
            });
        // Set the Value of the Select
        elSelect.value = RowColumn;
    
    // Add the Column Selector in a new cell in the new Row
        var elCell = document.createElement("td");
        elCell.appendChild( elSelect );
        elRow.appendChild( elCell );

    // Create the Operator Selector for Select
    if( SectionNumber == 1 ){
        var elOperator = document.createElement("select");
        elOperator.id = "Operator_" + RowUID;
        // Insert all Operators
            spr._op.forEach( function( thisOp ){
                var elOption = document.createElement("option");
                elOption.value = thisOp.op;
                elOption.text = thisOp.label;
                elOperator.appendChild( elOption );
            });
        // Set the Value of the Select
        elOperator.value = RowOperator;

        // Add this Selector in a new Cell
        var elCell = document.createElement("td");
        elCell.appendChild( elOperator );
        elRow.appendChild( elCell );
    }

    // Create the Direction Selector for Sort
    if( SectionNumber == 2 ){
        var elDirection = document.createElement("select");
        elDirection.id = "Direction_" + RowUID;
        // Insert all Operators
            spr._dir.forEach( function( thisDir ){
                var elOption = document.createElement("option");
                elOption.value = thisDir.direction;
                elOption.text = thisDir.label
                elDirection.appendChild( elOption );
            });
        // Set the Value of the Select
        elDirection.value = RowDirection;

        // Add this Selector in a new Cell
        var elCell = document.createElement("td");
        elCell.appendChild( elDirection );
        elRow.appendChild( elCell );
    }

    // Create a Textbox for Value/Label for Select and Show.
    if( SectionNumber == 1 || SectionNumber == 3 ){
        var elText = document.createElement("input");
        elText.setAttribute("type", "text");
        // Set the ID and Value
        switch( SectionNumber ){
            case 1:
                elText.id = "Value";
                elText.value = RowValue;
                break;
            case 3:
                elText.id = "Label";
                elText.value = RowLabel;
                break;
        }
        // Append RowUID to the ID
            elText.id = elText.id + "_" + RowUID;
        // Add this Text Input in a new Cell
        var elCell = document.createElement("td");
        elCell.appendChild( elText );
        elRow.appendChild( elCell );

    }
    
    
    // Add the set of options to this Row
        var elCell = document.createElement("td");
        // Source of Glyphs: https://www.svgrepo.com/collection/arrows-and-user-interface-2/
        elCell.innerHTML = `<!--
        <span class="tooltip pointer" onclick="spReportify.builderManipulateRow('up', '${elRow.id}');">
        <svg xmlns="http://www.w3.org/2000/svg" height="20" width="20" viewBox="0 0 512 512" xml:space="preserve"><path d="M476.9 216.5 263.5 3a10.6 10.6 0 0 0-15 0L35.2 216.5c-4 4.2-4 11 .2 15 4.1 4 10.7 4 14.8 0L245.3 36.4v465a10.7 10.7 0 0 0 21.3 0v-465l195.1 195c4.3 4.1 11 4 15-.1 4.1-4.2 4.1-10.7.2-14.8z"/></svg>
        <span class="tooltiptext">Move this row up</span></span>
        &nbsp;
        <span class="tooltip pointer" onclick="spReportify.builderManipulateRow('down', '${elRow.id}');">
        <svg xmlns="http://www.w3.org/2000/svg" height="20" width="20" viewBox="0 0 511.9 511.9" xml:space="preserve"><path d="M476.7 280.4a10.6 10.6 0 0 0-15 0L266.5 475.6V11a11 11 0 0 0-9-10.9c-6.7-1-12.3 4.2-12.3 10.6v465L50 280.3c-4.3-4-11-4-15 .2s-4 10.7 0 14.9l213.3 213.3a10.6 10.6 0 0 0 15 0l213.3-213.3c4.2-4 4.2-10.9 0-15z"/></svg>
        <span class="tooltiptext">Move this row down</span></span>
        &emsp;-->
        <span class="tooltip pointer" onclick="spReportify.builderManipulateRow('rm', '${elRow.id}');">
        <svg xmlns="http://www.w3.org/2000/svg" height="20" width="20" viewBox="0 0 511.9 511.9" xml:space="preserve"><path d="M271.2 255.9 509 18c4-4.2 4-11-.2-15S498-1 493.9 3L256.1 240.7 18.3 3c-4.3-4-11-4-15 .3-4 4.2-4 10.7 0 14.8L241 256 3.3 493.7c-4.3 4-4.4 10.8-.2 15a10.6 10.6 0 0 0 15 .2l.2-.2 237.8-237.8 237.7 237.8c4.3 4 11 4 15-.2s4-10.7 0-14.8L271.3 255.9z"/></svg>
        <span class="tooltiptext">Delete this row</span></span>
        `;
        elRow.appendChild( elCell );



    // Add the new row in the right table
        var TableName = "";
        switch( SectionNumber ){
            case 1:
                TableName = "BuilderFormSelectTableBody";
                break;
            case 2:
                TableName = "BuilderFormSortTableBody";
                break;
            case 3:
                TableName = "BuilderFormShowTableBody";
                break;
        }
        spr.logTrace("Destination of Row: " + TableName );
        document.getElementById(TableName).appendChild( elRow );
},

/**
 * builderManipulateRow
 * Manipulate a Row.
 * Action:  rm, up, down.
 */
builderManipulateRow: function( Action, RowId ){
    spr.logTrace("Action on Row: " + Action );
    spr.logTrace("Row ID: " + RowId );
    
    var elManipulate = document.getElementById( RowId );

    switch( Action ){
        case "rm": 
            elManipulate.parentNode.removeChild( elManipulate );
            break;
    }

},

builderSave: function(){
    spr.logInfo( "Save the Report ");

    // Build Select, Sort and Show Strings
    var QueryArray = [];
        // Select
        var ThisArray = [];
        var Rows = document.querySelectorAll('[id^="Select_"]');
        var ThisQueryArray = [];
        Rows.forEach( function( thisRow ){
            var thisRowUID = (thisRow.id).split("_")[1];
            var Column = document.getElementById("Column_" + thisRowUID).value;
            if( Column != "" ){
                var Operator = document.getElementById("Operator_" + thisRowUID).value;
                var Value = document.getElementById("Value_" + thisRowUID).value;
                // String
                ThisArray.push( ([ "R", Column, Operator, Value ]).join( spr.vm ) );
                // Query
                    var ThisOperator = spr._op.find( x => x.op == Operator );
                    console.log( `Details for Operator "${Operator}"` );
                    console.table( ThisOperator );
                    // Quote Text/Note Values
                    var ThisColumn = spReportifyData.builder.columns.find( x => x.name == Column );
                    console.log( `Details for Column "${Column}"` );
                    console.table( ThisColumn );
                    if( (["Text", "Note"]).includes( ThisColumn.type ) ){
                        Value = `'${Value}'`;
                    }
                    if( ThisOperator.syntax == "seq" ){
                        ThisQueryArray.push( `(${Column} ${Operator} ${Value})` );
                    }else{
                        ThisQueryArray.push( `(${Operator}(${Column},${Value})` );
                    }
            }
        });
        var StringSelect = ThisArray.join( "\n" );
        var QuerySelect = "";
            if( ThisQueryArray.length > 0 ){
                QuerySelect = "$filter=" + ThisQueryArray.join( " and " );
                QueryArray.push( QuerySelect );
            }

        spr.logTrace( 'SELECT' );
        spr.logTable( ThisArray );
        spr.logTrace( "String for the database" );
        spr.logTrace( StringSelect );
        spr.logTrace( "String for the query" );
        spr.logTrace( QuerySelect );

        // Sort
        var ThisArray = [];
        var ThisQueryArray = [];
        var Rows = document.querySelectorAll('[id^="Sort_"]');
    
        Rows.forEach( function( thisRow ){
            var thisRowUID = (thisRow.id).split("_")[1];
            var Column = document.getElementById("Column_" + thisRowUID).value;
            if( Column != "" ){
                var Direction = document.getElementById("Direction_" + thisRowUID).value;
                // String
                ThisArray.push( ([ "R", Column, Direction ]).join( spr.vm ) );
                // Query
                ThisQueryArray.push( `${Column} ${Direction}` );
            }
        });
        var StringSort = ThisArray.join( "\n" );
        var QuerySort = "";
            if( ThisQueryArray.length > 0 ){
                QuerySort = "$orderby=" + ThisQueryArray.join( "," );
                QueryArray.push( QuerySort );
            }

        spr.logTrace( 'SORT' );
        spr.logTable( ThisArray );
        spr.logTrace( "String for the database" );
        spr.logTrace( StringSort );
        spr.logTrace( "String for the query" );
        spr.logTrace( QuerySort );

        // Show
        var ThisArray = [];
        var ThisQueryArray = [];
        var Rows = document.querySelectorAll('[id^="Show_"]');
        var ShowContainsIdField = false;

        Rows.forEach( function( thisRow ){
            var thisRowUID = (thisRow.id).split("_")[1];
            var Column = document.getElementById("Column_" + thisRowUID).value;
            if( Column != "" ){
                var Label = document.getElementById("Label_" + thisRowUID).value;
                // Detect presence of ID field
                if( Column == "ID" ){
                    ShowContainsIdField = true;
                }
                // String
                ThisArray.push( ([ "R", Column, Label ]).join( spr.vm ) );
                // Query
                ThisQueryArray.push( Column );
            }
        });
        // Add ID if not already there
        if( ShowContainsIdField == false ){
            ThisQueryArray.push( "ID" );
        }
        var StringShow = ThisArray.join( "\n" );
        var QueryShow = "";
        if( ThisQueryArray.length > 0 ){
            QueryShow = "$select=" + ThisQueryArray.join( "," );
            QueryArray.push( QueryShow );
        }

        spr.logTrace( 'SHOW' );
        spr.logTable( ThisArray );
        spr.logTrace( "String for the database" );
        spr.logTrace( StringShow );
        spr.logTrace( "String for the query" );
        spr.logTrace( QueryShow );


        // Build the Query
        var Query = QueryArray.join( "&" );


        // Save to SharePoint Report List
            // Create or Retrieve the record
            if( spReportifyData.builder.mode == "edit" ){
                switch( (spReportifyData.config.reportListRefType).toLowerCase() ){
                    case "title":
                        this.ReportRecord = spReportifyData.sp.web.get_lists().getByTitle(spReportifyData.config.reportListRef).getItemById( spReportifyData.builder.report.id );
                        break;
                    case "guid":
                        this.ReportRecord = spReportifyData.sp.web.get_lists().getById(spReportifyData.config.reportListRef).getItemById( spReportifyData.builder.report.id );
                        break;
                }
                
            }else{
                this.NewReportRecord = new SP.ListItemCreationInformation();
                switch( (spReportifyData.config.reportListRefType).toLowerCase() ){
                    case "title":
                        this.ReportRecord = spReportifyData.sp.web.get_lists().getByTitle(spReportifyData.config.reportListRef).addItem( this.NewReportRecord );
                        break;
                    case "guid":
                        this.ReportRecord = spReportifyData.sp.web.get_lists().getById(spReportifyData.config.reportListRef).addItem( this.NewReportRecord );
                        break;
                }
                
            }

            // Set Values
            this.ReportRecord.set_item('Title', spReportifyData.builder.report.title );
            this.ReportRecord.set_item('ListId', spReportifyData.builder.list.id );
            this.ReportRecord.set_item('Description', document.getElementById("BuilderFormOptionDescription").value);

            this.ReportRecord.set_item('SelectEntries', StringSelect );
            this.ReportRecord.set_item('SortEntries', StringSort );
            this.ReportRecord.set_item('ShowEntries', StringShow );
            this.ReportRecord.set_item('Query', Query );

            // Update Record
            this.ReportRecord.update();

            // Send Query
            spReportifyData.sp.ctx.executeQueryAsync(
                Function.createDelegate( this, this.builderSave_Success ),
                Function.createDelegate( this, this.logSysError )
            );
},

builderSave_Success: function(){
    alert( 'Report Saved !');
},

runnerGetReport: function( Identifier ){
    spr.waitingShow( "Requesting the report definition..." );
    switch( (spReportifyData.config.reportListRefType).toLowerCase() ){
        case "title":
            spReportifyData.runner["api"] = spReportifyData.sp.web.get_lists().getByTitle( spReportifyData.config.reportListRef ).getItemById( spReportifyData.params.rpt );
            break;
        case "guid":
            spReportifyData.runner["api"] = spReportifyData.sp.web.get_lists().getById( spReportifyData.config.reportListRef ).getItemById( spReportifyData.params.rpt );
            break;
    }
    spReportifyData.sp.ctx.load( spReportifyData.runner.api );
    spReportifyData.sp.ctx.executeQueryAsync( 
        Function.createDelegate( this, spr.runnerParseReport ),
        Function.createDelegate( this, spr.logSysError )
    );
},

runnerParseReport: function(){
    spr.waitingShow( "Reading the report definition..." );
    spReportifyData.runner["report"] = {
        title: spReportifyData.runner.api.get_item("Title"),
        listId: spReportifyData.runner.api.get_item("ListId"),
        description: spReportifyData.runner.api.get_item("Description"),
        columns: spReportifyData.runner.api.get_item("ShowEntries"),
        query: spReportifyData.runner.api.get_item("Query")
    };
    spr.logTrace("Report Definition");
    spr.logTable( spReportifyData.runner.report );

    spr.runnerGetData();
},

runnerGetData: function(){
    spr.waitingShow( "Reading data..." );


    var AjaxOptions = {
        url: `${spReportifyData.config.url}/_api/web/Lists(guid'${spReportifyData.runner.report.listId}')/items?${spReportifyData.runner.report.query}` ,
        method: "GET",
        async: false,
        headers: {
            "accept":"application/json;odata=verbose",
            "content-type":"application/json;odata=verbose"
        }
    }

    // Send Request to Server
    $.ajax( AjaxOptions )
    .done( function ( data ){
        spReportifyData.runner["data"] = data.d.results;
        spr.runnerDrawReport();
        
    })
    .fail( function( jqXHR, textStatus, error ){
        spr.logError("Error Getting Report.");
        spr.logError( textStatus );
        spr.logError( error );
    })

},


runnerDrawReport: function(){

    document.getElementById("RunnerReportTitle").innerHTML = spReportifyData.runner.report.title;
    document.getElementById("RunnerRuntimeUser").innerHTML = spReportifyData.user.title + " on " + spReportifyData.site.title;

    // Decompose the structure of the Show Columns
    spReportifyData.runner["columns"] = [];
    spReportifyData.runner.report.columns.split("\n").forEach( function( thisLine ){
        var thisEntry = thisLine.split( spr.vm );
        spReportifyData.runner.columns.push({
            type: thisEntry[0],
            name: thisEntry[1],
            label: thisEntry[2]
        })
    });

    // Table Headers
    var elHeaders = document.createElement( "tr" );
    spReportifyData.runner.columns.forEach( function( thisColumn ){
        var elHeader = document.createElement( "td" );
        elHeader.innerHTML = thisColumn.label;
        elHeaders.appendChild( elHeader );
    });
    document.getElementById("RunnerReportTableHead").appendChild( elHeaders );

    // Details
    spReportifyData.runner.data.forEach( function( thisRecord ){
        var elRow = document.createElement( "tr" );
        elRow.id = "data_" + thisRecord["ID"];

            spReportifyData.runner.columns.forEach( function( thisColumn ){
                var elCell = document.createElement( "td" );
                elCell.id = thisColumn.name + "_" + thisRecord["ID"];
                elCell.innerHTML = thisRecord[thisColumn.name];
                elRow.appendChild( elCell );
            });
        document.getElementById("RunnerReportTableBody").appendChild( elRow );
    });

    spr.waitingHide();
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
        spr.logTrace("Parameters from URL");
        spr.logTable( spReportifyData.params );
        spr.stackRun();
    },


/**
 * waitingCreate
 * Add a waiting DIV in the Page to display when the script is working.
 */
    waitingCreate: function(){
        var ctlDivWaiting = document.createElement('div');
        ctlDivWaiting.id = "wip";
        ctlDivWaiting.classList.add("pleaseWait");
        ctlDivWaiting.innerHTML = `<div style="width:100%;height:100%;display:flex;align-items:flex-start;">
        <svg id="wipSpinner" width="100px" height="100px" viewBox="0 0 100 100" xmlnx="http://www.w3.org/2000/svg">
            <circle id="wipSpinnerAnime" cx="50" cy="50" r="45" />
        </svg><div><span id="wipTell"></span></div></div>`;
        document.body.appendChild( ctlDivWaiting );
        spr.show("wip");
        this.stackRun();
    },
/**
 * waitingShow
 * Display a message in the "Please Wait..." popup.
 */
    waitingShow: function( Message = "" ){
        document.getElementById( "wipTell" ).innerHTML = (Message == "" ? "One moment please..." : Message );
        spr.show( "wip" );
    },
/**
 * waitingHide
 * Hide the "Please Wait..." popup.
 */
    waitingHide: function(){
        spr.hide( "wip" );
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
        if( spReportifyData.config.logLevel > 0){
            console.log(`%c${Title}`, "color: #233E82;font-family:Tahoma;font-weight:bold;font-size:18px;");
        }
    },
/**
 * logMute
 * Display an entry in the Console, light gray.
 */
    logMute: function( Message ){
    console.log(`%c  ${Message}  `, "background: #D6E5E3;color:#04080F;font-family:Arial;");
},
/**
 * logError
 * If Log Level is at least "error".
 * Display an Error Console Entry.
 */
logError: function( Message ){
    if( spReportifyData.config.logLevel >= 4 ){
        console.log(`%c  ${Message}  `, "background: #d62828;color:#fff;font-family:Arial;font-weight:bold;");
    }
},
/**
 * logSysError
 * If Log Level is at least "error".
 * Display System-Error - To use for failures.
 */
logSysError: function( sender, args ){
    if( spReportifyData.config.logLevel >= 1 ){
        console.log(`%c ${args.get_message()}`, "background: #ff5d8f, color: black; font-family: Arial; font-weight: bold;");
    }
},
/**
 * logWarn
 * If Log Level is at least "warn".
 * Display a Warning Console Entry.
 */
logWarn: function( Message ){
    if( spReportifyData.config.logLevel >= 3 ){
        console.log(`%c  ${Message}  `, "background: #fca311;color:black;font-family:Arial;");
    }
},
/**
 * logInfo
 * If Log Level is at least "info".
 * Display an Information Console Entry.
 */
logInfo: function( Message ){
    if( spReportifyData.config.logLevel >= 2 ){
        console.log(`%c  ${Message}  `, "background: #cee5f2;color:black;font-family:Arial;");
    }
},
/**
 * logTrace
 * If Log Level is at least "trace".
 * Display an Information Console Entry.
 */
logTrace: function( Message ){
    if( spReportifyData.config.logLevel >= 1 ){
        console.log( Message );
    }
},
/**
 * logTable
 * Same level as Trace.
 * Display a DataTable in the Console.
 */
logTable: function( Data ){
    if( spReportifyData.config.logLevel >= 1 ){
        console.table( Data );
    }
},

/**
 * show
 * Show an Element.
 */
show: function( ElementId ){
    if( document.getElementById( ElementId ) ){
        document.getElementById( ElementId ).style.setProperty("display", "block");
    }else{
        spr.logError( `Requested to SHOW Element "${ElementId}" but this Element is NOT FOUND.` );
    }
},
/**
 * hide
 * Hide an Element.
 */
hide: function( ElementId ){
    if( document.getElementById( ElementId ) ){
        document.getElementById( ElementId ).style.setProperty("display", "none");
    }else{
        spr.logError( `Requested to HIDE Element "${ElementId}" but this Element is NOT FOUND.` );
    }
}


    


};

let spr = spReportify;

$(document).ready(function(){
    spReportify.init();
 });