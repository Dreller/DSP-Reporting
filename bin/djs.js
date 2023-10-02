// (c) DRELLER
// https://github.com/Dreller
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
// Core JS Functions
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
//
// General Variables
        var _ndx = 1;

//      Array of Filter Operations
        var _oper = [
            { op: "eq", label: "Equal", notation: "=", syntax: "seq" },
            { op: "ne", label: "Not equal", notation: "<>", syntax: "seq" },
            { op: "gt", label: "Greater than", notation: ">", syntax: "seq" },
            { op: "ge", label: "Greater than or equal", notation: ">=", syntax: "seq" },
            { op: "lt", label: "Less than", notation: "<", syntax: "seq" },
            { op: "le", label: "Less than or equal", notation: "<=", syntax: "seq" }
        ];

function HelloDreller(runtype){
    window["dreller"] = {
        client: {
            lang: ((window.navigator.language).split("-")[0]).toLowerCase(),
            param: {},
            timezone: Intl.DateTimeFormat().resolvedOptions().timeZone
        },
        queue: [],
        history: [],
        report: {},
        performance: {
            start: 0, 
            end: 0,
            processTime: 0,
            rows: 0
        },
        runtime: {
            mode: runtype.toLowerCase(),
            action: ""
        }
    }

// Running a Report
if( runtype == "run" ){
    d_StackAdd( d_ParseURL );
    d_StackAdd( d_GetCurrentUser );
    d_StackAdd( d_GetSite );
    d_StackAdd( d_LoadReport, null);
    d_StackAdd( d_BuildQueryUrl );
    d_StackAdd( d_RequestData );
}

// Building a Report
if( runtype == "build" ){
    d_StackAdd( d_ParseURL );
    d_StackAdd( d_GetCurrentUser );
    d_StackAdd( d_GetSite );
    d_StackAdd( d_BuildHeadings );
    d_StackAdd( d_LoadLists );
    d_StackAdd( d_InsertLists );

    d_StackAdd( d_Waiting, false );

}



// Execute the queue sequentially.
    d_StackRun();
}

function d_StackAdd(fctName, arguments = null){
    dreller.queue.push({fct: fctName, args: arguments});
}

function d_StackRun(){
    if( (dreller.queue).length == 0){ return; };
    var thisCommand = dreller.queue.shift();
    var thisCommandName = thisCommand.fct;
    var thisCommandArgs = thisCommand.args;
        if( thisCommandArgs == null ){
            thisCommandName();
        }else{
            thisCommandName(thisCommandArgs);
        }
}

function d_ParseURL(){
// Get URL Parameters and store them in dreller.client.param
        const urlParams = new URLSearchParams(window.location.search);
        const entries = urlParams.entries();
        for( const entry of entries ){
            dreller.client.param[(entry[0]).toLowerCase()] = entry[1];
        }
    d_StackRun();
}

function d_GetCurrentUser(){
    d_CallSP( "currentuser", null, function(response){
        dreller["user"] = {
            email: response.d.Email,
            name: response.d.Title,
            id: response.d.Id
        }
        d_StackRun();

    });
}

function d_GetSite(){
    d_CallSP( "title" , null, function(response){
        dreller["site"] = {
            title: response.d.Title,
            ctx: new SP.ClientContext( _URL )
        }
        dreller.site["api"] = dreller.site.ctx.get_web().get_lists().getByTitle( _REPORTLIST );
        d_StackRun();

    });
}

function d_LoadReport(args){
// Load an existing report.  args = {id: 0}
    if( args == null ){ var args = {id: dreller.client.param.rpt } }
    dreller.performance.start = performance.now();
    d_CallSP( "lists/getbytitle('" + _REPORTLIST +  "')/items(" + args.id + ")", null, function(response){
        var x1 = JSON.parse( response.d.Definition );
        var x2 = {
            title: response.d.Title,
            id: response.d.ID,
            list_guid: response.d.List
        };
        dreller["report"] = {
            ...x2,
            ...x1
        }
        d_StackRun();
    });
}

function d_BuildQueryUrl(){
// https://www.odata.org/documentation/odata-version-2-0/uri-conventions/#FilterSystemQueryOption

    // Where
        var aWhere = [];
        dreller.report.select.forEach( function(thisSelect, thisIndex){
            var thisOp = _oper.filter( x => (x.op).toLowerCase() == (thisSelect.op).toLowerCase() )[0];
            if( typeof thisOp === "undefined" ){
                console.error( `Oops!  Operator not handled: ${thisSelect.op}` );
            }else{

                console.log( thisOp );

                var thisVal = `'${thisSelect.value}'`;
                    if( thisSelect.type == "Number" ){
                        thisVal = `${thisSelect.value}`;
                    }
                    
                if( thisOp.syntax == "seq" ){
                    aWhere.push( `(${thisSelect.col} ${thisOp.op} ${thisVal})` );
                }else{
                    aWhere.push( `(${thisOp.op}(${thisSelect.col}, ${thisVal}))` );
                }

            }
            
        });

    // Sort
        var aSort = [];
        dreller.report.sort.forEach( function( thisSort, thisIndex ){
            aSort.push( `${thisSort.col} ${thisSort.order}` );
        });

    // Show
        var aShow = [];
        dreller.report.show.forEach( function( thisShow, thisIndex ){
            aShow.push(thisShow.col);
        });

    // Build the Query
        var queryString = `$format=json&$filter=(${ aWhere.join(" and ") })&$orderby=${ aSort.join(",") }&$select=${ aShow.join(",") }`;
        dreller.report["odata"] = {
            string: queryString,
            where: aWhere,
            sort: aSort,
            show: aShow
        }

    console.log( dreller.report.odata );
    d_StackRun();

}

function d_RequestData(){
    d_CallSP( "Lists(guid'" + dreller.report.list_guid + "')/items?" + dreller.report.odata.string, 
    { method: "GET" },
    function(response){
        d_BuildTable(response.d.results);
        d_BuildHeadings();
    })
}

function d_BuildTable(data){
    // Insert Table Headings
    dreller.report.show.forEach( function( thisShow, thisIndex ){
        $("#reportContainerHead").append(`<th data-column="${thisShow.col}" data-position="${thisShow.i}" data-type="${thisShow.type}">${thisShow.title}</th>`);
    } );

    // Insert Rows
    data.forEach( function( thisRow, thisIndex){
        $("#reportContainerBody").append(`<tr data-uri="${thisRow.__metadata.uri}" data-sptype="${thisRow.__metadata.type}" data-guid="${thisRow.__metadata.id}" data-id="${thisRow.ID}">`);
        dreller.report.show.forEach( function( thisShow, thisIndex ){
            $("#reportContainerBody").append(d_ProcessCell(thisRow.ID, thisRow[thisShow.col], thisShow));
        });
        $("#reportContainerBody").append(`</tr>`);
    });

    // Count Rows
    dreller.performance.rows = data.length;
}

function d_ProcessCell(id, value, col){
    
    var cellStyle = "";
    var cellValue;
    var cellFormat = "";

    switch(col.type){
        case "Number":
            cellFormat = "Number";
            cellStyle = "text-align:right;"
            cellValue = (value).toLocaleString(dreller.client.lang);
            break;
        
        default:
            cellFormat = "Default";
            cellStyle = "text-align: left;";
            cellValue = value;
    }
    
    return `<td data-itemid="${id}" data-column="${col.col}" data-format="${cellFormat}" style="${cellStyle}">${cellValue}</td>`;
}

function d_BuildHeadings(){
    if( dreller.runtime.mode == "run" ){
        // Report Name
            $("#reportName").html(dreller.report.title);
        // Report Description
            $("#runInfo").html(dreller.report.description);

        // Runtime Info
            dreller.performance.end = performance.now();
            dreller.performance.processTime = (dreller.performance.end - dreller.performance.start);
            $("#runInfo").append(`<p>Report run by ${dreller.user.name} on ${dreller.site.title}<br>
            ${dreller.performance.rows} rows processed in ${Math.round(dreller.performance.processTime)} milliseconds, on ${d_PrettyNow()}.</p>`);
    }

    if( dreller.runtime.mode == "build" ){
        // User Info
            $("#headUserLogon").html(`Logged as ${dreller.user.name} on ${dreller.site.title}`);
    }

    d_StackRun();
    
}



function d_LoadList(args){

}

function d_LoadColumn(args){

} 

/**
 * Load Lists for the current Site.
 */
function d_LoadLists(){
    d_CallSP("lists?$select=Id,Title,Hidden,BaseType,ItemCount,Fields&$expand=Fields", {}, function(response){
        dreller.site["lists"] = [];
        response.d.results.forEach( function( thisEntry, thisIndex) {
            // Establish conditions if we should keep the list or not.
                // Don't keep if list is "Hidden".
                    if( thisEntry.Hidden == true ){ 
                        console.log(`(d_LoadLists) List "${thisEntry.Title}" is Hidden."`);
                        return;
                    }
                // ...
            // At this point, we keep the list
                (dreller.site.lists).push( thisEntry );
        });
    });

    d_StackRun();
}

/**
 * Insert Lists in Select
 */
function d_InsertLists(){
    (dreller.site.lists).forEach( function( thisList, thisIndex ){
        $("#lstDatasource").append(`<option value="${thisList.Id}">${thisList.Title} - ${(thisList.BaseType=="1"?"Library":"List")} - ${thisList.ItemCount} record(s).</option>`);
    });
    d_StackRun();
}

/**
 * Get Reports for a List
 */
function d_GetReports(){
    d_WaitingMessage( "Getting Reports for this Source..." );
    d_Waiting( true );
    d_CallSP( "lists/getbytitle('" + _REPORTLIST +  "')/items?$select=Id,Title,Definition,List&$filter=List eq '" + $("#lstDatasource").val() + "'", null, function(response){
        if( response.d.results.length == 0 ){
            $("#tabExisting").hide();
            $("#tabCreate").hide();
            $('#SelectReportExisting').hide();
            $('#SelectReportCreate').show();
            dreller.runtime.action = "create";
        }else{
            $("#tabExisting").hide();
            $("#tabCreate").show();
            $('#SelectReportExisting').show();
            $('#SelectReportCreate').hide()
            dreller.runtime.action = "edit";
            dreller.site["reports"] = response.d.results;
            $("#lstReport").empty();
            $("#lstReport").append(`<option value="">Choose a report...</option>`);
            (response.d.results).forEach( function( thisReport, thisIndex ){
                $("#lstReport").append(`<option value="${thisReport.Id}">${thisReport.Title}</option>`);
            });
        }
        d_Waiting( false );
    });
}

/**
 * Build the Report Editor
 */
function d_BuildEditor(){
    var xList = (dreller.site.lists).filter( x => x.Id == $("#lstDatasource").val() )[0];
    var xReport = (dreller.site.reports).filter( y => y.Id == $("#lstReport").val())[0];

    dreller["editor"] = {
        list: xList,
        columns: xList.Fields.results,
        report: xReport,
        reportName: (dreller.runtime.action=="create" ? $("#txtNewReportName").val():xReport.Title)
    }

    // Build the Select Template for Columns, so we build it only once.
        var sSelectTemplate = `<select id="selectColumn">`;
        dreller.editor.columns.forEach( function(thisColumn){
            sSelectTemplate += `<option value="${thisColumn.InternalName}">${thisColumn.Title} (${thisColumn.TypeAsString})</option>`;
            $("#dataDictBody").append(`<tr><td>${thisColumn.Title}</td><td>${thisColumn.TypeAsString}</td><td>${thisColumn.Description}</td></tr>`);
        });
        sSelectTemplate += `</select>`;
        

    // Build the Select Template for Operators
        var sSelectOperator = `<select id="selectOperator">`;
            _oper.forEach( function( thisOper){
                sSelectOperator += `<option value="${thisOper.op}">${thisOper.label}</option>`;
            });
        sSelectOperator += "</select>";

    // Build the Select Template for Sort Direction
        var sSelectOrder = `<select id="selectDirection"><option value="asc">Ascending (A-Z)</option><option value="desc">Descending (Z-A)</option></select>`;

    // Store Templates
        dreller.editor["templates"] = {
            selColumn: sSelectTemplate,
            selOperator: sSelectOperator,
            selDirection: sSelectOrder
        };

    // Hide Selectors
        $("#SelectDatasource").hide();
        $("#SelectReport").hide();

    // Display what we are doing and on what
        $("#headSelectionDone").html( (dreller.runtime.action == "create" ? `Creating` : `Editing`) + ` Report "${dreller.editor.reportName}" on Datasource "${dreller.editor.list.Title}".` );
        $(".drellerSelectionDone").show();

    // Set the Report's 3S:  Select / Sort / Show

        // Select
            // Add the new blank line
                $("#selectTBody").append(`<tr id="selectTemplate" style="display:none;" data-rowtype="template" data-category="select"><td>
                    ${d_BuilderControl({type:"select-col"})}
                </td><td>
                    ${d_BuilderControl({type:"select-oper"})}
                </td><td>
                    <input type="text" id="selectValue">
                </td>
                <td>
                    <span id="optionDelete" class="icon" data-tooltip="Remove this row" onclick="d_BuilderRemRow(this.parentElement.parentElement);">&#59153;</span>
                </td></tr>`);

        // Sort
                $("#sortTBody").append(`<tr id="sortTemplate" style="display:none;" data-rowtype="template" data-category="sort"><td>
                    ${d_BuilderControl({type:"select-col"})}
                </td><td>
                    ${d_BuilderControl({type:"select-sort"})}
                </td><td>
                <span id="optionDelete" class="icon" data-tooltip="Remove this row" onclick="d_BuilderRemRow(this.parentElement.parentElement);">&#59153;</span>
            </td></tr>`);

        // Show
                $("#showTBody").append(`<tr id="showTemplate" style="display:none;" data-rowtype="template" data-category="show"><td>
                    ${d_BuilderControl({type:"select-col"})}
                </td><td>
                    <input type="text" id="showLabel">
                </td><td>
                <span id="optionDelete" class="icon" data-tooltip="Remove this row" onclick="d_BuilderRemRow(this.parentElement.parentElement);">&#59153;</span>
            </td></tr>`);

    // Parse the Report Definition
        dreller.editor.report["defn"]  = JSON.parse( dreller.editor.report.Definition );

    // Show Selection Criteria
        dreller.editor.report.defn.select.forEach( function( thisSelect ){
            d_BuilderAddItem({
                type: "select",
                blank: false,
                data: thisSelect
            });
        } );
    // Show Sort Order
        dreller.editor.report.defn.sort.forEach( function( thisSort ){
            d_BuilderAddItem({
                type: "sort",
                blank: false,
                data: thisSort
            });
        } );
    // Show Columns to show
        dreller.editor.report.defn.show.forEach( function( thisShow ){
            d_BuilderAddItem({
                type: "show",
                blank: false,
                data: thisShow
            });
        } );

    // Set the Report Options



    // Show the Editor
    $("#ReportBuilderSection").show();

}

function d_BuilderRemRow( RowElement ){
    console.log("ID: " + RowElement.id);
    RowElement.remove();
}

/**
 * Return a Report Item
 * @param {} Options - {
 *                          type: select|sort|show
 *                          blank: true|false   (to add a copy of the template)
 *                          data: [record]
 *                     }
 */
function d_BuilderAddItem( Options ){
    console.log( Options );
    // Get the Template Line
        var templateRow = document.getElementById( Options.type + "Template" );
    // Get the Table
        var containerTable = document.getElementById( Options.type + "TBody" );
    // Create the clone
        var clonedRow = templateRow.cloneNode( true );
    // Adjust the ID and the data-type
        clonedRow.id = Math.random();
        clonedRow.dataset.rowtype = "data";
    // Make the clone visible
        clonedRow.style.display = "";

console.log( clonedRow );

        // If the blank parameter is NOT true, we have values to set in the new line
        if( Options.hasOwnProperty( "blank" ) && Options.blank != true ){

            clonedRow.querySelector('[id="selectColumn"]').value = Options.data.col;

            switch( Options.type ){
                case "select":
                    clonedRow.querySelector('[id="selectOperator"]').value = (Options.data.op).toLowerCase();
                    clonedRow.querySelector('[id="selectValue"]').value = Options.data.value;
                    break;
                case "sort":
                    clonedRow.querySelector('[id="selectDirection"]').value = (Options.data.order).toLowerCase();
                    break;
                case "show":
                    clonedRow.querySelector('[id="showLabel"]').value = Options.data.title;
                    break;
            } 
        }

    // Add the clone in the table
        containerTable.appendChild( clonedRow );
}


/**
 * Return a Control for the Report Builder
 * @param {} Options - {
 *                       type: select-col|select-sort|select-oper|text
 *                       id: ID to set in the control
 *                     }
 */
function d_BuilderControl(Options){
    var WipControl;
    switch( Options.type ){
        case "select-col":
            WipControl = dreller.editor.templates.selColumn;
            break;
        case "select-oper":
            WipControl = dreller.editor.templates.selOperator;
            break;
        case "select-sort":
            WipControl = dreller.editor.templates.selDirection;
            break;
        case "text":
            WipControl = `<input type="text" id="%ID%" />`;
            break;
    }

    if( Options.hasOwnProperty("id") ){
        WipControl = WipControl.replaceAll( "%ID%", Options.id );
    }

    return WipControl;
}

function d_BuilderSaveReport(){
    // TODO: Prompt Validations...


    // Build the Report Definition
    var oSelect = [];   var iSelect = 1;
    var oSort = [];     var iSort = 1;
    var oShow = [];     var iShow = 1;
    document.querySelectorAll('[data-rowtype="data"]').forEach(function(thisRow){

        var thisColumn = thisRow.querySelector('[id="selectColumn"]').value;
        var thisColumnType = (dreller.editor.columns.filter( x => x.InternalName == thisColumn )[0]).TypeAsString;
        console.log("Processing column: '" + thisColumn + "', type: " + thisColumnType);

        switch( thisRow.dataset.category ){
            case "select":
                oSelect.push({
                    i: iSelect,
                    col: thisColumn,
                    op: thisRow.querySelector('[id="selectOperator"]').value,
                    type: thisColumnType,
                    value: thisRow.querySelector('[id="selectValue"]').value
                });
                iSelect++;
                break;
            case "sort":
                oSort.push({
                    i: iSort,
                    col: thisColumn,
                    order: thisRow.querySelector('[id="selectDirection"]').value,
                    type: thisColumnType
                });
                iSort++;
                break;
            case "show":
                oShow.push({
                    i: iShow,
                    col: thisColumn,
                    title: thisRow.querySelector('[id="showLabel"]').value,
                    type: thisColumnType
                });
                iShow++;
                break;
        }
    });

    var oDefn = {
        description: "",
        select: oSelect,
        sort: oSort,
        show: oShow
    }

    console.log( oDefn );

    // SSave Report to SharePoint
        if( dreller.runtime.action == "edit" && dreller.editor.report.Id > 0 ){
            this.oListItem = dreller.site.api.getItemById(dreller.editor.report.Id);
        }else{
            var oCreateInfo = new SP.ListItemCreationInformation();
            this.oListItem = dreller.site.api.addItem( oCreateInfo );
        }
    
    this.oListItem.set_item('Definition', JSON.stringify(oDefn));
    this.oListItem.set_item('Title', dreller.editor.reportName);
    this.oListItem.set_item('List', dreller.editor.list.Id );
    this.oListItem.update();

    dreller.site.ctx.executeQueryAsync(
        Function.createDelegate( this, this.d_BuilderSaveReport_Success ),
        Function.createDelegate( this, this.d_BuilderSaveReport_Failed )
    );

}

function d_BuilderSaveReport_Success(){
    alert( "Report" + (dreller.runtime.action == "edit" ? " Updated":" Created") + "." );
    dreller.runtime.action = "edit";
}

function d_BuilderSaveReport_Failed(sender, args){
    alert('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}


function d_PrettyNow(){
    var raw = new Date();
    
    var aMonths = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
    if( dreller.client.lang == "fr" ){
        aMonths[1] = "FEV";
        aMonths[3] = "AVR";
        aMonths[4] = "MAI";
        aMonths[7] = "AOU";
    }
    var aParts = [];
        // Day
            aParts.push( ("00" + raw.getDate() ).substr(-2) );
        // Month
            aParts.push( aMonths[raw.getMonth()] );
        // Year
            aParts.push( raw.getFullYear() );
        // Time
            aParts.push( ( "00" + raw.getHours() ).substr(-2) + ":" + ( "00" + raw.getMinutes() ).substr(-2) + ":" + ( "00" + raw.getSeconds() ).substr(-2) );
        // Timezone
            aParts.push( "(" + dreller.client.timezone + ")" );

    return aParts.join(" ");
}

function d_Waiting(toggle){
    if( toggle ){
        $("#waiting").show();
    }else{
        $("#waiting").hide();
    }
}

function d_WaitingMessage(text){
    $("#waitingInfo").html(text);
}

function d_CallSP(endpoint, args, callback){
    let basics = {
        url: _URL + "/_api/web/" + endpoint,
        method: "GET",
        async: false,
        headers: {
            "accept": "application/json;odata=verbose",
            "content-type": "application/json;odata=verbose"
        }
    }
    let params = {
        ...basics,
        ...args
    }
    // Add Digest if available
    if( typeof dreller.digest !== undefined && dreller.digest != "" ){
        params.headers["X-RequestDigest"] = dreller.digest;
    }

    console.log( params.method + " to " + params.url );

    $.ajax(params)
    .done( function ( data ) {
        callback( data );
    })
    .fail( function( data ) {
        console.error("> SharePoint API Error");
        if( typeof data.responseJSON.error.message.value !== undefined ){
            console.error( data.responseJSON.error.message.value );
        }
    })
    .always( function( data ) {
        console.dir( data );
        dreller.history["api" + _ndx] = data;
        _ndx++;
    });
}







/*
 * Modal
 *
 * Pico.css - https://picocss.com
 * Copyright 2019-2023 - Licensed under MIT
 */

// Config
const isOpenClass = "modal-is-open";
const openingClass = "modal-is-opening";
const closingClass = "modal-is-closing";
const animationDuration = 400; // ms
let visibleModal = null;

// Toggle modal
const toggleModal = (event) => {
  event.preventDefault();
  const modal = document.getElementById(event.currentTarget.getAttribute("data-target"));
  typeof modal != "undefined" && modal != null && isModalOpen(modal)
    ? closeModal(modal)
    : openModal(modal);
};

// Is modal open
const isModalOpen = (modal) => {
  return modal.hasAttribute("open") && modal.getAttribute("open") != "false" ? true : false;
};

// Open modal
const openModal = (modal) => {
  if (isScrollbarVisible()) {
    document.documentElement.style.setProperty("--scrollbar-width", `${getScrollbarWidth()}px`);
  }
  document.documentElement.classList.add(isOpenClass, openingClass);
  setTimeout(() => {
    visibleModal = modal;
    document.documentElement.classList.remove(openingClass);
  }, animationDuration);
  modal.setAttribute("open", true);
};

// Close modal
const closeModal = (modal) => {
  visibleModal = null;
  document.documentElement.classList.add(closingClass);
  setTimeout(() => {
    document.documentElement.classList.remove(closingClass, isOpenClass);
    document.documentElement.style.removeProperty("--scrollbar-width");
    modal.removeAttribute("open");
  }, animationDuration);
};

// Close with a click outside
document.addEventListener("click", (event) => {
  if (visibleModal != null) {
    const modalContent = visibleModal.querySelector("article");
    const isClickInside = modalContent.contains(event.target);
    !isClickInside && closeModal(visibleModal);
  }
});

// Close with Esc key
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && visibleModal != null) {
    closeModal(visibleModal);
  }
});

// Get scrollbar width
const getScrollbarWidth = () => {
  // Creating invisible container
  const outer = document.createElement("div");
  outer.style.visibility = "hidden";
  outer.style.overflow = "scroll"; // forcing scrollbar to appear
  outer.style.msOverflowStyle = "scrollbar"; // needed for WinJS apps
  document.body.appendChild(outer);

  // Creating inner element and placing it in the container
  const inner = document.createElement("div");
  outer.appendChild(inner);

  // Calculating difference between container's full width and the child width
  const scrollbarWidth = outer.offsetWidth - inner.offsetWidth;

  // Removing temporary elements from the DOM
  outer.parentNode.removeChild(outer);

  return scrollbarWidth;
};

// Is scrollbar visible
const isScrollbarVisible = () => {
  return document.body.scrollHeight > screen.height;
};