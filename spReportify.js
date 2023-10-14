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
        builder: {}         // Work Data for the Report Builder.
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
            spReportifyData.builder.mode = "edit";
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
    var listId = document.getElementById("BuilderFormControlDatasource").value
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
builderToggleMode: function(TargetMode = 1){
    var NewMode = ( TargetMode == 1 || typeof TargetMode == "undefined" || typeof spReportifyData.builder.mode == "undefined" ? "add" : "edit" );
    if( NewMode == spReportifyData.builder.mode ){ return; }
    spr.logTitle( "Changing Builder Mode" );
    console.log( `Current Mode: "${spReportifyData.builder.mode}"` );
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
    console.log( `New Mode: "${spReportifyData.builder.mode}"` );
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

    spr.logTitle(`Report to ${spReportifyData.builder.mode}`);
    console.table( spReportifyData.builder.report );

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
    console.log('Row UID: ' + RowUID ); 

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
        elRow.id = RowUID;

    // Create the Column Selector
        var elSelect = document.createElement("select");
        elSelect.id = "Column";
        // Insert Columns in the Select
            spReportifyData.builder.columns.forEach( function( thisColumn ){
                var elOption = document.createElement("option");
                elOption.value = thisColumn.name;
                elOption.text = thisColumn.title;
                elSelect.appendChild( elOption );
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
        elOperator.id = "Operator";
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
        elDirection.id = "Direction";
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
        // Add this Text Input in a new Cell
        var elCell = document.createElement("td");
        elCell.appendChild( elText );
        elRow.appendChild( elCell );

    }
    
    
    // Add the set of options to this Row
        var elCell = document.createElement("td");
        // Source of Glyphs: https://www.svgrepo.com/collection/arrows-and-user-interface-2/
        elCell.innerHTML = `<!--
        <span class="tooltip pointer" onclick="spReportify.builderManipulateRow('up', '${RowUID}');">
        <svg xmlns="http://www.w3.org/2000/svg" height="20" width="20" viewBox="0 0 512 512" xml:space="preserve"><path d="M476.9 216.5 263.5 3a10.6 10.6 0 0 0-15 0L35.2 216.5c-4 4.2-4 11 .2 15 4.1 4 10.7 4 14.8 0L245.3 36.4v465a10.7 10.7 0 0 0 21.3 0v-465l195.1 195c4.3 4.1 11 4 15-.1 4.1-4.2 4.1-10.7.2-14.8z"/></svg>
        <span class="tooltiptext">Move this row up</span></span>
        &nbsp;
        <span class="tooltip pointer" onclick="spReportify.builderManipulateRow('down', '${RowUID}');">
        <svg xmlns="http://www.w3.org/2000/svg" height="20" width="20" viewBox="0 0 511.9 511.9" xml:space="preserve"><path d="M476.7 280.4a10.6 10.6 0 0 0-15 0L266.5 475.6V11a11 11 0 0 0-9-10.9c-6.7-1-12.3 4.2-12.3 10.6v465L50 280.3c-4.3-4-11-4-15 .2s-4 10.7 0 14.9l213.3 213.3a10.6 10.6 0 0 0 15 0l213.3-213.3c4.2-4 4.2-10.9 0-15z"/></svg>
        <span class="tooltiptext">Move this row down</span></span>
        &emsp;-->
        <span class="tooltip pointer" onclick="spReportify.builderManipulateRow('rm', '${RowUID}');">
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
        spr.logMute("Destination of Row: " + TableName );
        document.getElementById(TableName).appendChild( elRow );
},

/**
 * builderManipulateRow
 * Manipulate a Row.
 * Action:  rm, up, down.
 */
builderManipulateRow: function( Action, RowId ){
    spr.logTitle("Action on Row: " + Action );
    spr.logMute("Row ID: " + RowId );
    
    var elManipulate = document.getElementById( RowId );

    switch( Action ){
        case "rm": 
            elManipulate.parentNode.removeChild( elManipulate );
            break;
    }

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
},
/**
 * logError
 * Display an entry in the Console, bold error.
 */
logError: function( Message ){
    console.log(`%c  ${Message}  `, "background: #E81123;color:#fff;font-family:Arial;font-weight:bold;");
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