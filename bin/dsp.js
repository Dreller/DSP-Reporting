// (c) DRELLER
// https://github.com/Dreller
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
// JS Functions to interact with SharePoint Online.
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
// Good reference here: https://hevodata.com/learn/sharepoint-api/#s
// 

function dSpInit(){
    window["dreller"]["sp"] = {
        api: dreller.config.rootSiteUrl + "/_api/web/",
        ctx: new SP.ClientContext( dreller.config.rootSiteUrl )
    }
    dreller.sp["lib"] = dreller.sp.ctx.get_web().get_lists().getByTitle( dreller.config.reportListName );
    dStackAdd("dSpGetSiteName");
    dStackAdd("dSpGetCurrentUser");
    dStackAdd("dSpGetLists");
    dStackAdd("dSpGetReports");
    dStackAdd("dSpDrawLists", "dFormListName");
    dStackRun();
}

function dSpRunInit(){
    window["dreller"]["sp"] = {
        api: dreller.config.rootSiteUrl + "/_api/web/",
        ctx: new SP.ClientContext( dreller.config.rootSiteUrl )
    }
    dreller.sp["lib"] = dreller.sp.ctx.get_web().get_lists().getByTitle( dreller.config.reportListName );

    dStackAdd("dSpGetSiteName");
    dStackAdd("dSpGetCurrentUser");
    dStackAdd("dSpGetSingleReport")
    dStackRun();
}


function dSpCall(urlSuffix, args, callback){
    let basics = {
        url: dreller.sp.api + urlSuffix,
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
    $.ajax(params)
    .done( function ( data ) {
        console.log( data );
        callback( data );
    })
    .fail( function( data ) {
        console.error("> SharePoint API Error");
        console.log( data );
    });
}


function dSpGetLists(){
    dSpCall( "lists", {}, 
        function( thisData ){
            dreller["_lists"] = [];
            thisData.d.results.forEach( function( thisEntry ){
                if( thisEntry.Hidden != true ){
                    (dreller._lists).push( thisEntry );
                }
            } );
            dStackRun();
        }
    )
}

function dSpGetReports(){
    dSpCall( "lists/getbytitle('" + dreller.config.reportListName +  "')/items", {}, 
        function( thisData ){
            dreller["_reports"] = [];
            thisData.d.results.forEach( function( thisEntry ){
                (dreller._reports).push( thisEntry );
            } );
            dStackRun();
        }
    )
}

function dSpDrawLists(targetSelect){
    dFormSelectSetOptions( targetSelect, dreller._lists, "Title","Id" );
    dStackRun();
}

function dSpSelectList(){
    dreller["list"] = {
        id: $("#dFormListName").val(),
        record: (dreller._lists).filter( x=> x.Id == $("#dFormListName").val() )[0]
    }
    dSpDrawReports();
}

function dSpDrawReports(){
    dFormSelectSetOptions( "dFormReportName", (dreller._reports).filter( x => x.List == $("#dFormListName").val() ), "Title","Id" );
}

function dSpDrawFields(){
    (dreller.columns).forEach( function( thisColumn ){
        $("select[data-source='col']").append(`<option value="${thisColumn.InternalName}">${thisColumn.Title}</option>`);
    });
    dStackRun();
}

function dSpGetSiteName(){
    dSpCall( "title", {}, 
        function( thisData ){
            dreller["site"] = {
                title: thisData.d.Title
            }
        }
    );
    dStackRun();
}

function dSpGetCurrentUser(){
    dSpCall( "currentuser", {},
        function(thisData){
            console.log( thisData );
            dreller["usr"] = {
                name: thisData.d.Title,
                email: thisData.d.Email,
                id: thisData.d.Id
            }
            $("#userInfo").html(`Logged as ${dreller.usr.name} on ${dreller.site.title} site.`);
            dStackRun();
        }
    )
}

function dSpEditReport(){
    window["dreller"]["report"] = {};
    dreller.report = {
        list: $( "#dFormListName" ).val(),
        record: ( dreller._reports ).find( x => x.Id == $( "#dFormReportName" ).val()[0] ),
        id: $( "#dFormReportName" ).val(),
        name: ""
    }
    dreller.report.name = dreller.report.record.Title;

    dStackAdd("dSpGetListColumns");
    dStackAdd("dSpDrawReport");
    dStackRun();
}

function dSpDrawReport(){
    $("#dForm").hide();
    $("#userInfo").html( $("#userInfo").html() + " &emsp; Data Source: " + dreller.list.record.Title + " &emsp; Report: " + dreller.report.record.Title );

    var oDefn = JSON.parse(dreller.report.record.Definition);
    console.dir( oDefn );

        (oDefn.select).forEach(function(thisItem){
            dAddRow("Select", thisItem);
        });
        (oDefn.sort).forEach(function(thisItem){
            dAddRow("Sort", thisItem);
        });
        (oDefn.show).forEach(function(thisItem){
            dAddRow("Show", thisItem);
        });
    // Set Report Options
        $("#rptOptDescription").val(oDefn.description);
        $("#rptOptName").val(dreller.report.name);
}

function dSpGetListColumns(){
    dreller["columns"] = [];
    dreller["_columns"] = [];
    dreller["_list"] = dreller.sp.ctx.get_web().get_lists().getById(dreller.list.id);
    dreller._columns = (dreller._list).get_fields();
    dreller.sp.ctx.load(dreller._columns);
    dreller.sp.ctx.executeQueryAsync(
        dSpGetListColumns_POST, dSpServerFailed
    );
}

function dSpGetListColumns_POST(){
    var oEnum = dreller._columns.getEnumerator();
    while( oEnum.moveNext() ){
        var oCol = oEnum.get_current();
        (dreller.columns).push({
            Title: oCol.get_title(),
            StaticName: oCol.get_staticName(),
            InternalName: oCol.get_internalName(),
            TypeString: oCol.get_typeAsString(),
            TypeDisplayName: oCol.get_typeDisplayName(),
            TypeDescription: oCol.get_typeShortDescription(),
            Description: oCol.get_description(),
            Hidden: oCol.get_hidden(),
            Id: oCol.get_id(),
            Indexed: oCol.get_indexed(),
            Group: oCol.get_group()
        });
        console.log(oCol.get_typeAsString());
    }
    (dreller.columns).sort(function(a,b){
        var x = (a["Title"]).toLowerCase();
        var y = (b["Title"]).toLowerCase();
        if( x == y){return 0;}
        if( x > y){return 1;}else{return -1;}
    });
    dSpDrawFields();
}

function dSpCreateReport(){
    dreller["report"] = {};
    dreller.report = {
        record: {},
        id: -1,
        //name: ($( "#dFormReportNameEntry" ).val()).trim().split(" ").join(".").toUpperCase(),
        name: ($( "#dFormReportNameEntry" ).val()).trim(),
        list: dreller.list.id
    }

    var itemCreateInfo = new SP.ListItemCreationInformation();
    this.oListItem = dreller.sp.lib.addItem(itemCreateInfo);
    oListItem.set_item('Title', dreller.report.name);
    oListItem.set_item('List', dreller.report.list);
    oListItem.update();

    dreller.sp.ctx.load(oListItem);
    dreller.sp.ctx.executeQueryAsync(
        Function.createDelegate(this,  this.dSpCreateReport_POST),
        Function.createDelegate(this, this.dSpServerFailed)
    );

}

function dSpCreateReport_POST(){
    var itemId = oListItem.get_id();
    dreller.report.id = itemId;
    dreller.report.record = {
        Title: this.oListItem.get_item('Title'),
        ID: itemId,
        GUID: "",
        List: this.oListItem.get_item('List'),
        Definition: ""
    };
    (dreller.reports).push( dreller.report.record );
    dStackAdd("dSpGetListColumns");
    dStackAdd("dSpDrawReport");
    dStackRun();
}


function dSpServerFailed(sender, args){
    console.error(args.get_stackTrace());
    alert('Request Failed!  ' + args.get_message() + '\n' + args.get_stackTrace());
    }

function dSpSaveReport(){
    // Process Select
        var oSelect = [];
        var ndx = 1;
        $("#tableSelect > tbody > tr").each( function(thisNdx, thisLine){
            var sSelCol = $(this).find("select:first").val();
            if( (sSelCol + "").length == 0 ){return;}
            var sSelOp = $(this).find("select:last").val();
            var sSelVal = $(this).find("input").val();
            var sSelTyp = (dreller.columns.filter( x => x.InternalName == sSelCol )[0]).TypeString;

            var oWip = {
                i: ndx,
                col: sSelCol,
                op: sSelOp,
                value: sSelVal,
                type: sSelTyp
            }

            oSelect.push(oWip);
            ndx++;

        });
    // Process Sort
        var oSort = [];
        var ndx = 1;
        $("#tableSort > tbody > tr").each( function(thisNdx, thisLine){
            var sSorCol = $(this).find("select:first").val();
            if( (sSorCol + "").length == 0 ){return;}
            var sSorOrd = $(this).find("select:last").val();

            var oWip = {
                i: ndx,
                col: sSorCol,
                order: sSorOrd
            }

            oSort.push(oWip);
            ndx++;
        });
    // Process Show
        var oShow = [];
        var ndx = 1;
        $("#tableShow > tbody > tr").each( function(thisNdx, thisLine){
            var sShoCol = $(this).find("select:first").val();
            if( (sShoCol + "").length == 0 ){return;}
            var sShoTi = $(this).find("input").val();
            var sShoTy = (dreller.columns).filter( x => x.InternalName == $(this).find("select:first").val() )[0].TypeString;

            var oWip = {
                i: ndx,
                col: sShoCol,
                title: sShoTi,
                type: sShoTy
            }

            oShow.push(oWip);
            ndx++;
        });

    // Combine Definition and catch other values to save.
        var oDefn = {
            select: oSelect,
            sort: oSort,
            show: oShow,
            description: ($("#rptOptDescription").val()).trim()
        }
        var sTitle = ($("#rptOptName").val()).trim();

    // Get Record from SharePoint
        this.oListItem = dreller.sp.lib.getItemById(dreller.report.id);
        this.oListItem.set_item('Definition', JSON.stringify(oDefn));
        this.oListItem.set_item('Title', sTitle);
        this.oListItem.update();

        dreller.sp.ctx.executeQueryAsync(
            Function.createDelegate(this, this.dSpSaveReport_POST),
            Function.createDelegate(this.dSpServerFailed)
        );
}

function dSpSaveReport_POST(){
    alert("Saved");
}

function dSpGetSingleReport(){
    dreller["report"] = {
        id: parseInt(dreller.params.rpt),
        name: ""
    }
    console.log(dreller.report);
    this.oListItem = dreller.sp.lib.getItemById(dreller.report.id);
    dreller.sp.ctx.load(oListItem);
    dreller.sp.ctx.executeQueryAsync(
        Function.createDelegate(this, this.dSpGetSingleReport_POST),
        Function.createDelegate(this.dSpServerFailed)
    );
}

function dSpGetSingleReport_POST(){
    dreller.report["record"] = {
            Definition: oListItem.get_item('Definition')
    };
    dreller.report.name = oListItem.get_item('Title')

    dreller["_list"] = dreller.sp.ctx.get_web().get_lists().getById(oListItem.get_item('List'));
    dSpBuildCAML();
}

function dSpBuildCAML(){
    dreller["caml"] = {
        query: "",
        _api: new SP.CamlQuery()
    };

    var oDefn = JSON.parse(dreller.report.record.Definition);

    var camlWhere = "";
    (oDefn.select).forEach( function( thisItem ){
        camlWhere += `<${thisItem.op}><FieldRef Name='${thisItem.col}'/><Value Type='${thisItem.type}'>${thisItem.value}</Value></${thisItem.op}>`;
    });
    
    var camlOrderBy = "";
    (oDefn.sort).forEach( function( thisItem ){
        camlOrderBy += `<FieldRef Ascending='${ ( (thisItem.order == "asc") ? "TRUE":"FALSE" ) }' Name='${thisItem.col}' />`;
    });

    var camlFields = "";
    (oDefn.show).forEach( function( thisItem ){
        camlFields += `<FieldRef Name='${thisItem.col}'/>`;
    });

    dreller.caml.query = `<View><Query><Where>${camlWhere}</Where>`;
    if( camlOrderBy != "" ){
        dreller.caml.query += `<OrderBy>${camlOrderBy}</OrderBy>`;
    }

    dreller.caml.query += `</Query></View>`;

    console.log(dreller.caml.query);
    
    (dreller.caml._api).set_viewXml(dreller.caml.query);
    dreller.caml["_res"] = dreller._list.getItems(dreller.caml._api);
    dreller.sp.ctx.load(dreller.caml._res);
    dreller.sp.ctx.executeQueryAsync(Function.createDelegate(this, dSpBuildCAML_POST), Function.createDelegate(this, dSpServerFailed));

}

function dSpBuildCAML_POST(){
    var oDefn = JSON.parse(dreller.report.record.Definition);
    $("#reportContainerHead").append("<tr>");
    (oDefn.show).forEach( function( thisItem ){
        $("#reportContainerHead").append(`<th>${thisItem.title}</th>`);
    });

    $("#reportContainerHead").append("</tr>"); 
    var e = dreller.caml._res.getEnumerator();
    while( e.moveNext() ){
        var el = e.get_current();
        $("#reportContainerBody").append(`<tr>`);
            (oDefn.show).forEach( function( thisItem ){
                $("#reportContainerBody").append(`<td>${el.get_item(thisItem.col)}</td>`);
            });
        $("#reportContainerBody").append(`</tr>`);
    }


    $("#reportName").html(dreller.report.name);
    $("#runInfo").html(`Generated by ${dreller.usr.name}, insert date-time here.<br>Source: ${dreller.site.tile} site.`);
    // Add Report Description
        if( oDefn.hasOwnProperty("description") ){
            $("#runInfo").append(`<p>${oDefn.description}</p>`);
        }

}