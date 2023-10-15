<!DOCTYPE html>
<head>
    <title>dSP Report Builder</title>
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
            <h3>Options</h3>
            ...
            
        </section>
    


</section>  <!-- Builder Form -->
    <!-- REPORT BUILDER -->
    <section id="ReportBuilderSection" style="display:none;">


        <dialog id="data-dict">
            <article>
                <header>
                    <a href="#close" aria-label="Close" class="close" data-target="data-dict" onclick="toggleModal(event)"></a>
                    Data Dictionary
                </header>
                <table>
                    <thead>
                        <tr>
                            <th>Column</th>
                            <th>Type</th>
                            <th>Description</th>
                        </tr>
                    </thead>
                    <tbody id="dataDictBody">

                    </tbody>
                </table>
            </article>
          </dialog>

        <!-- Select -->
            <details open>
                <summary role="button" class="secondary">
                    Select
                </summary>
                <p>
                    <table id="selectTable">
                        <thead>
                            <tr>
                                <th>
                                    Column
                                </th>
                                <th>
                                    Operator
                                </th>
                                <th>
                                    Value
                                </th>
                                <th>
                                    Options
                                </th>
                            </tr>
                        </thead>
                        <tbody id="selectTBody"></tbody>
                    </table>
                </p>
                <div style="display: inline-block;padding: 12px 0px;">
                    <a href="#!" onclick="d_BuilderAddItem({type:'select'})">Add row</a>
                </div>
            </details>

        <!-- Sort -->
            <details open>
                <summary role="button" class="secondary">
                    Sort
                </summary>
                <p>
                    <table id="sortTable">
                        <thead>
                            <tr>
                                <th>
                                    Column
                                </th>
                                <th>
                                    Direction
                                </th>
                                <th>
                                    Options
                                </th>
                            </tr>
                        </thead>
                        <tbody id="sortTBody"></tbody>
                    </table>
                </p>
                <div style="display: inline-block;padding: 12px 0px;">
                    <a href="#!" onclick="d_BuilderAddItem({type:'sort'})">Add row</a>
                </div>
            </details>

        <!-- Show -->
            <details open>
                <summary role="button" class="secondary">
                    Show
                </summary>
                <p>
                    <table id="showTable">
                        <thead>
                            <tr>
                                <th>
                                    Column
                                </th>
                                <th>
                                    Column Title
                                </th>
                                <th>
                                    Options
                                </th>
                            </tr>
                        </thead>
                        <tbody id="showTBody"></tbody>
                    </table>
                </p>
                <div style="display: inline-block;padding: 12px 0px;">
                    <a href="#!" onclick="d_BuilderAddItem({type:'show'})">Add row</a>
                </div>
            </details>

        <!-- Report options -->
            <details>
                <summary role="button" class="contrast">
                    Report Options
                </summary>
                <p>
                    Additional options goes here.
                </p>
            </details>

        <!-- Report URL Parameters -->
            <details>
                <summary role="button" class="contrast">
                    URL Parameters
                </summary>
                <p>
                    Mappings goes here.  It will be a feature to use URL Parameters in <em>Select</em>.
                </p>
            </details>



        <!-- Save Button -->
        <p>
            <a href="#!" role="button" class="contrast" onclick="d_BuilderSaveReport();">Save report</a>
        </p>

    </section>
</main>



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