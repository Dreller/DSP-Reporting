// (c) DRELLER
// https://github.com/Dreller
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
// Core JS Functions
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
//
// Global Variables
    let spinnerEnabled;

function dInit(){
    // Initialize d
        window["dreller"] = {
            queue: []
        };

    // Get URL Parameters
        dreller["params"] = [];
        const queryString = window.location.search;
        const urlParams = new URLSearchParams(queryString);
        const entries = urlParams.entries();
        for( const entry of entries ){
            console.log(entry);
            (dreller.params).push({
                key: entry[0],
                value: entry[1]
            });
        }

    // Load Settings
        dStackAdd( "dJSON_Load", {binFile: "d.json",storage: "config"});
    // Initialize SP Code
        dStackAdd( "dSpInit" );
    // Start Execution
        dStackRun();
}

function dInitReport(){
    // Initialize d
        window["dreller"] = {
            queue: []
        };

    // Get URL Parameters
        dreller["params"] = {};
        const queryString = window.location.search;
        const urlParams = new URLSearchParams(queryString);
        const entries = urlParams.entries();
        for( const entry of entries ){
            dreller.params[entry[0]] = entry[1];
        }

    // Load Settings
        dStackAdd( "dJSON_Load", {binFile: "d.json",storage: "config"});
    // Initialize SP Code
        dStackAdd( "dSpRunInit" );
    // Start Execution
        dStackRun();
}

function dSpinner(){

}

function dJSON_Load(args){
    console.log("Executing dLoadJSON...");
    console.dir(args);
    if( dJSON_checkRequiredParams(args, ["binFile", "storage"]) == false ){
        alert(`Unable to load JSON File, see the Console for details.`);
        return false;
    }
    // Load the actual file into the storage variable
        let explodedURL = (window.location.href).split("/");
        explodedURL.pop();
        let fileURL = explodedURL.join("/") + "/bin/" + args.binFile;
        $.ajax({
            url: fileURL,
            method: "GET",
            dataType: "json",
            contentType: "application/json",
            success: function(data){
                if( typeof data === "object" ){
                    dreller[args.storage] = data;
                } else {
                    dreller[args.storage] = JSON.parse(data);
                }
                dStackRun();
            }
        });

}

function dJSON_checkRequiredParams(args, requiredParams){
// args: object to verify
// requiredParams: array of property names to check
    let missingParams = [];
    requiredParams.forEach(function(param){
        if( !args.hasOwnProperty(param)){
            missingParams.push(param);
        }
    });
    if( missingParams.length > 0 ){
        console.group("Missing Parameters");
        console.error(missingParams.length + " missing parameter(s) on " + requiredParams.length + " expected.");
        console.groupEnd();
        return false;
    }else{
        return true;
    }
}

function dStackAdd(fct, args = {}){
    console.log("Stack: " + fct + ", Args: " + JSON.stringify( args ));
    (dreller.queue).push({
            name: fct, 
            params: JSON.stringify(args)
        });
}
function dStackRun(){
    if( (dreller.queue).length > 0 ){
        
        let thisEntry = (dreller.queue).shift();
        console.log( thisEntry );
        let thisFct = thisEntry.name;
        console.log("Run from Stack: " + thisFct + ", Args: " + thisEntry.params );
        window[thisFct]( JSON.parse( thisEntry.params ) );
    }
}

function dFormSelectSetOptions(selectId, dataSource, displayName, valueName){
    $("#" + selectId).empty();
    $("#" + selectId).append("<option value=''>Select...</option>");
    dataSource.forEach( function( thisEntry ){
        $("#" + selectId).append(`<option value=${thisEntry[valueName]}>${thisEntry[displayName]}</option>`)
    });
}


function dAddRow(table, args = {}){
console.log("Add Row in: " + table);
console.log(args);
    var rowTemplate = $('#dEditor' + table + "_0").clone();
    var rowCount = $('#table' + table + " > tbody").find('tr').length;
    var newSuffix = parseInt(rowCount) + 1;
    var newPrefix = "dEditor" + table;
    var newId = newPrefix + "_" + newSuffix;

    $(rowTemplate).attr('id', newId);

    if( args == {} ){
        $('#table' + table + ' tbody').append(rowTemplate);
    }else{
        //$('#table' + table + ' tbody').prepend(rowTemplate);
        $(rowTemplate).insertBefore("#dEditor" + table + "_0");
    }
    
    // Set IDs
    $("#" + newId).find("select, input").each(function(ndx){
        $(this).attr('id', ($(this).attr('data-section') + $(this).attr('data-source') + newSuffix) );
        $(this).attr('data-ndx', newSuffix);
    });


    $("#" + newId).find("input[data-ndx=" + newSuffix + "]").each( function(thisNdx){
        if( args.hasOwnProperty( $(this).attr('data-source') ) ){
            $(this).val(args[$(this).attr('data-source')]);
        }else{
            $(this).val("");
        }
    });
    $("#" + newId).find("select[data-ndx=" + newSuffix + "]").each( function(thisNdx){
        if( args.hasOwnProperty( $(this).attr('data-source') ) ){
            $(this).val(args[$(this).attr('data-source')]);
        }else{
            $(this).val("");
        }
    });
}
