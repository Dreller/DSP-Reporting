<!DOCTYPE html>
<head>
    <title>spReportify &#10149; Home</title>
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
    <style>
        #HomeCards{
            display: table;
            border-collapse: separate;
            border-spacing: 10px 10px;
        }
        .HomeCard{
            display: table-cell;
            box-shadow: 0 4px 8px 0 rgba(0,0,0,0.2);
            transition: 0.3s;
            border-radius: 5px;
            cursor:pointer;
        }
        .HomeCard:hover{
            box-shadow: 0 8px 16px 0 rgba(0,0,0,0.2);
        }
        .HomeCardContainer{
            padding: 2px 16px;
        }
    </style>

</head>
<body>
   <!-- Header -->
   <header>
        <span>SP Reportify</span>
        <span>A Reporting Solution for your data in SharePoint!</span>
   </header>

   <!-- Section with Cards to get to SPR Tools -->
   <section id="HomeCards">

        <div class="HomeCard">
            <div class="HomeCardContainer" onclick="location.href=spReportifyData.config.pageBuilder;">
                <h4>Report Builder</h4>
                <p>Create and edit reports.</p>
            </div>
        </div>

        <div class="HomeCard" onclick="spReportify.homeGetReports();">
            <div class="HomeCardContainer">
                <h4>Report Directory</h4>
                <p>View a list of all reports.</p>
            </div>
        </div>

        <div class="HomeCard">
            <div class="HomeCardContainer">
                <h4>Dictionary Explorer</h4>
                <p>Explore your Data Dictionary.</p>
            </div>
        </div>
   </section>

   <!-- Report Directory -->
    <section id="HomeReports" style="display:none;">
        <div style="width: 100%; text-align: right; margin-bottom: 5px;">
            <span class="button" onclick="spReportify.homeBackHome();">
                Close
            </span>
        </div>
        <table id="HomeReportsTable">
            <thead>
                <tr>
                    <th>List</th>
                    <th>Report Name</th>
                    <th>Description</th>
                    <th>Options</th>
                </tr>
            </thead>
            <tbody id="HomeReportsBody">

            </tbody>
        </table>
    </section>

</body>