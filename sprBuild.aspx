<!DOCTYPE html>
<head>
    <title>spReportify ● Builder</title>
    <link rel="icon" type="image/x-icon" href="../_layouts/15/images/favicon.ico?rev=23">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <script src="bin/jquery-3.7.0.js"></script>

    <script src="https://ajax.aspnetcdn.com/ajax/4.0/MicrosoftAjax.js" type="text/javascript"></script> 
    <script type="text/javascript" src="../_layouts/15/sp.runtime.js"></script>
    <script type="text/javascript" src="../_layouts/15/sp.js"></script>

    <script src="bin/dconf.js"></script>
    <script src="spReportify.js"></script>

    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Roboto&display=swap" rel="stylesheet">

    <link rel="stylesheet" href="spReportify.css">

</head>
<body>
   <!-- Header -->
   <header>
        <span>SP Reportify</span>
        <span>A Reporting Solution for your data in SharePoint!</span>
   </header>

<!-- First form: Select Datasource + Report -->
<section id="BuilderIdentifyReport">
<!-- Select the Datasource -->
    <section id="BuilderFormSectionDatasource">
        <label for="BuilderFormControlDatasource">
            Select a Datasource:
        </label>
        <select id="BuilderFormControlDatasource">
            
        </select>
    </section>


<!-- Choose the Action to do: Edit or Create a Report -->
    <section id="BuilderFormSectionAction" style="display:none;">
        <div class="BuilderFormWrapperAction">
            <input type="radio" name="BuilderFormControlAction" id="BuilderFormControlActionChoiceEdit" onclick="spReportify.builderToggleMode(0);" checked>
            <input type="radio" name="BuilderFormControlAction" id="BuilderFormControlActionChoiceCreate" onclick="spReportify.builderToggleMode(1);">
            <label for="BuilderFormControlActionChoiceEdit" class="radioOption choiceEdit">
                <div class="BuilderFormControlActionDot"></div>
                <span>Edit a Report</span>
            </label>
            <label for="BuilderFormControlActionChoiceCreate" class="radioOption choiceCreate">
                <div class="BuilderFormControlActionDot"></div>
                <span>Create a Report</span>
            </label>
        </div>
    </section>

<!-- Select the Report to Edit -->
    <section id="BuilderFormSectionReportPicker" style="display:none;">
        <label for="BuilderFormControlReportPicker">
            Choose the report to edit:
        </label>
        <select id="BuilderFormControlReportPicker">
        </select>
    </section>

<!-- Create a new Report -->
    <section id="BuilderFormSectionReportNaming" style="display:none;">
        <label for="BuilderFormControlReportNaming">
            Give a name for your new report:
        </label>
        <input type="text" id="BuilderFormControlReportNaming" autocomplete="off" onkeyup="spReportify.builderValidateReportName();" />
        <span id="BuilderFormAlertReportNaming_AlreadyUsed" style="display:none;">
            This Report Name is already used for this Datasource!  Please use another name.
        </span>
        
    </section>
<!-- Continue button (Load/Create Report) -->
    <section id="BuilderFormSectionLoadCreateReport" style="display:none;">
        <a href="#!" id="BuilderFormControlButtonLoadCreate" role="button" class="button">Continue</a>
    </section>

</section> <!-- End of BuilderIdentifyReport-->


<!-- Read-Only View of the Datasource and Report Selection -->
<section id="BuilderReportIdentity" style="display: none;">
    <table>
        <thead>
            <tr>
                <th>
                    Datasource
                </th>
                <th>
                    Report
                </th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td id="BuilderReportIdentityDatasource">

                </td>
                <td id="BuilderReportIdentityReportName">
                    
                </td>
            </tr>
        </tbody>
    </table>
</section>

<!-- Builder Form -->
    <section id="BuilderForm" style="display:none;">

    <!-- Toolbar -->
        <section id="BuilderToolbar" class="toolbar">
            <ul class="toolbar">
            <li onclick="spReportify.builderCloseReport();">Close</li>
            <li data-modal="BuilderDataDictionary">Dictionary</li>
            <li onclick="spReportify.builderSave();">Save</li>
            <li onclick="abc">Save & Run</li>
            </ul>
        </section>

    <!-- Select -->
        <section id="BuilderFormSelect">
            <h3>Selection</h3>
            <table id="BuilderFormSelectTable">
                <thead id="BuilderFormSelectTableHead">
                    <tr>
                        <th>Column</th>
                        <th>Operator</th>
                        <th>Value</th>
                        <th>Options</th>
                    </tr>
                </thead>
                <tbody id="BuilderFormSelectTableBody">

                </tbody>
            </table>
            <a href="#!" id="BuilderFormSelectAddRow" role="button" class="button" onclick="spReportify.builderDrawRow(1);">Add Row</a>
        </section>

    <!-- Sort -->
        <section id="BuilderFormSort">
            <h3>Sort</h3>
            <table id="BuilderFormSortTable">
                <thead id="BuilderFormSortTableHead">
                    <tr>
                        <th>Column</th>
                        <th>Direction</th>
                        <th>Options</th>
                    </tr>
                </thead>
                <tbody id="BuilderFormSortTableBody">
                    
                </tbody>
            </table>
            <a href="#!" id="BuilderFormSelectAddRow" role="button" class="button" onclick="spReportify.builderDrawRow(2);">Add Row</a>
        </section>

    <!-- Show -->
        <section id="BuilderFormShow">
            <h3>Show</h3>
            <table id="BuilderFormShowTable">
                <thead id="BuilderFormShowTableHead">
                    <tr>
                        <th>Column</th>
                        <th>Header</th>
                        <th>Options</th>
                    </tr>
                </thead>
                <tbody id="BuilderFormShowTableBody">
                    
                </tbody>
            </table>
            <a href="#!" id="BuilderFormSelectAddRow" role="button" class="button" onclick="spReportify.builderDrawRow(3);">Add Row</a>
        </section>

    <!-- Options -->
        <section id="BuilderFormOptions">
            <h3>Other Settings and Options</h3>
            <table>
                <thead>
                    <tr>
                        <th>
                            Option
                        </th>
                        <th>
                            Description
                        </th>
                        <th>
                            Value
                        </th>
                    </tr>
                </thead>
                <tbody>
                    <!-- Report Description -->
                    <tr>
                        <td>
                            Report Description
                        </td>
                        <td>
                            Description displayed when a user is running the report.  It can be used to describe what the user is looking at.
                        </td>
                        <td>
                            <textarea id="BuilderFormOptionDescription">

                            </textarea>
                        </td>
                    </tr>
                    <!-- Report Max Items to returns per page -->
                    <tr>
                        <td>
                            Items by page
                        </td>
                        <td>
                            Number of items to display per page.
                        </td>
                        <td>
                            <input type="number" id="BuilderFormOptionBatchSize" />
                        </td>
                    </tr>
                    <!-- .. -->
                </tbody>
            </table>
            
        </section>
    


</section>  <!-- Builder Form -->



<!-- Data Dictionary -->
<div id="BuilderDataDictionary" class="modal">
    <div class="modal-bg modal-exit"></div>
    <div class="modal-container">
        <div style="width: 100%;">
            <h2 style="float:left;">Data Dictionary</h2>
            <button style="float:right;cursor:pointer;" class="modal-close modal-exit">X</button>
        </div>
        <table id="BuilderDictionaryTable">
            <thead id="BuilderDictionaryTableHead">
                <tr>
                    <th>
                        Column Title
                    </th>
                    <th>
                        Static Name
                    </th>
                    <th>
                        Data Type
                    </th>
                    <th>
                        Description
                    </th>
                </tr>
            </thead>
            <tbody id="BuilderDictionaryTableBody">

            </tbody>
        </table>
    </div>
    
</div>



</body>