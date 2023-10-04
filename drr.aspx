<!DOCTYPE html>
<head>
    <title>dSP Report</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <script src="bin/jquery-3.7.0.js"></script>

    <script src="https://ajax.aspnetcdn.com/ajax/4.0/MicrosoftAjax.js" type="text/javascript"></script> 
    <script type="text/javascript" src="../_layouts/15/sp.runtime.js"></script>
    <script type="text/javascript" src="../_layouts/15/sp.js"></script>
    <script src="bin/dconf.js"></script>
    <script src="bin/d.js"></script>
    <script src="bin/djs.js"></script>

    <style>

        *{
            font-family: Helvetica, Arial, sans-serif;
        }
        
        h1{
            color: #0f4c81; 
        }
        
        table{
            border-collapse: separate;
            border-spacing: 0;
            border-radius: 10px;
            width:100%;
        }
        th, td{
            padding: 7px 20px;
        }
        td{
            border: solid 1px gainsboro;
        }
        th{
            padding: 7px 20px;
            border-top: solid 2px gainsboro;
            border-bottom: solid 2px gainsboro;
            border-right: solid 1px gainsboro;
            border-left: solid 1px gainsboro;
        }
        thead tr, tr:hover{
            background-color: #f8f8f8;
        }
        th:first-child{
            border-top-left-radius:7px;
            border-left: solid 2px gainsboro;
            border-right: solid 1px gainsboro;
        }
        th:last-child{
            border-top-right-radius:7px;
            border-right: solid 2px gainsboro;
            border-left: solid 1px gainsboro;
        }
        
        [data-type="Currency"]{
            text-align: right;
        }
        td[data-type="Counter"]{
            text-align: left;
            font-family: Consola, monospace;
        }
        [data-type="Text"]{
            text-align: left;
        }
        
        caption{
            caption-side:bottom;
            text-align:left;
            padding-top: 1em;
            font-size: 10pt;
            color: #6e6e6e;
        }
        
        
        
        </style>

    <script>
        $(document).ready(function(){
            HelloDreller("run");
        });
    </script>
</head>
<body style="max-width: 800px; margin: 20px auto; padding: 0 20px;">

<h1 id="reportName"></h1>
<span id="headUserLogon"></span>
<table id="reportContainer">
    <caption id="runInfo"></caption>
    <thead id="reportContainerHead">
    </thead>
    <tbody id="reportContainerBody">
    </tbody>
</table>

</body>