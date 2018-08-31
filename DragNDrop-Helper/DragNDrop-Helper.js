/**
 * DragNDrop-Helper
 * This simple script helps you to process a simple task for one or a bunch of files.
 * Instead of repeatly dealing with command lines for the same task but with different files.
 * Works too for setting up a pipe.
 * Use it, modify it, whatever. :-)
 * @author Adrian Kauz
 * @version 1.0
 */
var sScriptName = "DragNDrop-Helper";
var oArgs = WScript.Arguments;
var oCom;

var oSettings = {
                    "script": {
                                "path"         : "",
                                "waitOnReturn" : true,
                                "showWindow"  : 0,
                                "runFile"      : "run.cmd" 
                              },
                    "pipe":   [
                    			  // Set up here your commandline.
	                    		  {
	                                "path"      : "",
	                                "exe"       : "",
	                                "parameter" : ""
	                              }
                    		  ]
				}

main();

 /*
================
main()
================
*/
function main()
{
	oCom = new ComObjects();

    if (!oCom.loadAllObjects()) {
        WScript.echo("Can't load all Com-Objects!");
        return;
    }

	if (oArgs.length === 0) {
		showExclamationBox("Can't do work without arguments (Files)!");
		return;
	}

	preparePaths();
	
	if (!checkPaths()) {
		return;
	}

	for (var x = 0; x < oArgs.length; x++) {
		if (!oCom.FSO.fileExists(oArgs(x))) {
			showExclamationBox("File \"" + oArgs(x) + "\" not found!");
			return;
		}

		var sPathToCmd = oSettings["script"]["path"] + "\\" + oSettings["script"]["runFile"];
		var oFile = oCom.FSO.CreateTextFile(sPathToCmd, true);
		oFile.WriteLine(getFullCommandLine(oArgs(x)));
		oFile.Close();

		// Now do work!
		oCom.Shell.Run(sPathToCmd, oSettings["script"]["showWindow"], oSettings["script"]["waitOnReturn"]);

		oCom.FSO.GetFile(sPathToCmd).Delete();
	}
}


/*
================
ComObjects()
================
*/
function ComObjects()
{
    this.Shell = null;
    this.FSO = null;

    this.loadAllObjects = function() {
        try {
            this.Shell = new ActiveXObject("WScript.Shell");
            this.FSO = new ActiveXObject("Scripting.FileSystemObject");
            return true;
        } catch(ex) {
            return false;
        }
    };
}


/*
================
getCommandLine()
================
*/
function getFullCommandLine(sFileNameWithPath)
{
	var sCmd = "";

	for (var x = 0; x < oSettings["pipe"].length; x++) {
		if (x > 0) {
			sCmd += " | ";
		}

		sCmd += getCommand(oSettings["pipe"][x]);
	}

	// Prepare Input-File
	sCmd = sCmd.replace("$INPUT$", sFileNameWithPath);

	// Prepare Output-File
	var arrSplit = sFileNameWithPath.split(".");
	arrSplit.pop();
	sCmd = sCmd.replace("$OUTPUT$", arrSplit.join("."));
	
	return sCmd;
}


/*
================
getCommand()
================
*/
function getCommand(oPipeElement)
{
	var sCmdPipeElement = "\"" + oPipeElement["path"] + "\\" + oPipeElement["exe"] + "\" " + oPipeElement["parameter"];

	return sCmdPipeElement;
}


/*
================
preparePaths()
================
*/
function preparePaths()
{
	oSettings["script"]["path"] = getPath(WScript.ScriptFullName);

	for (var x = 0; x < oSettings["pipe"].length; x++) {
		if (oSettings["pipe"][x]["path"].length === 0) {
			oSettings["pipe"][x]["path"] = oSettings["script"]["path"];
		}
	}
}
	

/*
================
getScriptPath()
================
*/
function getPath(sPath)
{
	arrSplit = sPath.split("\\");
	arrSplit.pop();
	return arrSplit.join("\\");
}


/*
================
getFullPathFromSettings()
================
*/
function getFullPath(oElement)
{
	sPath = oElement["path"] + "\\" + oElement["exe"];
	return sPath.replace(/\\/g, "\\\\");
}


/*
================
checkPaths()
================
*/
function checkPaths()
{
	for (var x = 0; x < oSettings["pipe"].length; x++) {
		if (!oCom.FSO.fileExists(getFullPath(oSettings["pipe"][x]))) {
			showExclamationBox("\"" + oSettings["pipe"][x]["exe"] + "\" not found!");
			
			return false;
		}
	}

	return true;
}


/*
================
showExclamationBox()
================
*/
function showExclamationBox(sMessage)
{
    oCom.Shell.popup(sMessage, 0, sScriptName, 48 );
}


/*
================
showInfoBox()
================
*/
function showInfoBox(sMessage)
{
    oCom.Shell.popup(sMessage, 0, sScriptName, 64 );
}
