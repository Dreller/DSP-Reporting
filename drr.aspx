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

    <script>
        $(document).ready(function(){
            HelloDreller("run");
        });
    </script>
</head>
<body style="max-width: 800px; margin: 20px auto; padding: 0 20px;">
    <h2 id="reportName"></h2>
    <hr>
    <blockquote><cite id="runInfo">Loading informations...</cite></blockquote>
<p>&nbsp;</p>
<table id="reportContainer">
    <thead id="reportContainerHead">
    </thead>
    <tbody id="reportContainerBody">
    </tbody>
</table>


</body>