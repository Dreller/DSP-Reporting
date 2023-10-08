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
    <script src="bin/d.js"></script>
    <script src="bin/djs.js"></script>

    <script src="spReportify.js"></script>

    <link rel="stylesheet" href="bin/pico.css">
    <style>
        .icon{
            font-family: 'Segoe MDL2 Assets';
            cursor:pointer;
        }

    </style>

    <script>
    
        $(document).ready(function(){
           // dInit();
            HelloDreller("build");
        });
    </script>
</head>
<body>
    <main class="container">

    <!-- Header -->
        <div>
            <hgroup>
                <h2>Report Builder for SharePoint</h2>
                <h3 id="headUserLogon"></h3>
            </hgroup>
            <footer class="drellerSelectionDone" style="padding: 15px; background: #f9fafb; display:none;" id="headSelectionDone"></footer>
            <hr class="drellerSelectionDone" style="display:none;">
        </div>

    <!-- One Moment Please -->
        <div id="waiting" style="position: fixed;width: 100%; top: 15%; height: 350px; left: 0; justify-content: center; display: flex;">
            <article id="waitingInfo" aria-busy="true">
                One moment please...
            </article>
        </div>

    <!-- Selector: Datasource (list) -->
        <div id="SelectDatasource">
            <label for="lstDatasource">Datasource:</label>
            <select id="lstDatasource">
                <option value="">Choose an option...</option>
            </select>
        </div>

    <!-- Selector: Select Existing Report or create new one -->
        <div id="SelectReport">
            <div style="width:100%;text-align:right;">
                <a id="tabExisting" href="#" style="display:none;" onclick="$('#SelectReportExisting').show();$('#SelectReportCreate').hide();$('#tabExisting').hide();$('#tabCreate').show();dreller.runtime.action = 'edit';">
                    Edit an existing report...
                </a>
                <a id="tabCreate" href="#" style="" onclick="$('#SelectReportExisting').hide();$('#SelectReportCreate').show();$('#tabExisting').show();$('#tabCreate').hide();dreller.runtime.action = 'create';">
                    Create a new report...
                </a>
            </div>
            <div id="SelectReportExisting" style="padding-top: 2rem;">
                <label for="lstReport">
                    Report to edit:
                </label>
                <select id="lstReport">
                    <option value="">Choose an option...</option>
                </select>
            </div>
            <div id="SelectReportCreate" style="display:none;padding-top: 2rem;">
                <label for="txtNewReportName">
                    Give your report a name:
                </label>
                <input type="text" id="txtNewReportName" />
            </div>
            <div style="width: 100%; text-align:right;">
                <a href="#" role="button" class="primary" onclick="d_BuildEditor()">Ok</a>
            </div>
        </div>


    <!-- REPORT BUILDER -->
    <section id="ReportBuilderSection" style="display:none;">
        <div style="display: inline-block;padding: 12px 0px;">
            <a href="#!" onclick="$('details').prop('open', true);">Expand all</a>&nbsp;|&nbsp;
            <a href="#!" onclick="$('details').prop('open', false);">Collapse all</a>&nbsp;|&nbsp;
            <a href="#!" onclick="toggleModal(event)" data-target="data-dict">Dictionary</a>
        </div>

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

</body>