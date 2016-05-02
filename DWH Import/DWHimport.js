!INC EAScriptLib.JScript-CSV
!INC EAScriptLib.JScript-Dialog
!INC Local Scripts.EAConstants-JScript


/*
 * This code has been included from the default Project Browser template.
 * If you wish to modify this template, it is located in the Config\Script Templates
 * directory of your EA install path.   
 * 
 * Script Name:
 * Author:
 * Purpose:
 * Date:
 */

var execMode = 0;
var selectedPackage as EA.Package;

/*
 * Project Browser Script main function
 */
function OnProjectBrowserScript()
{
	// Get the type of element selected in the Project Browser
	var treeSelectedType = Repository.GetTreeSelectedItemType();
	
	// Handling Code: Uncomment any types you wish this script to support
	// NOTE: You can toggle comments on multiple lines that are currently
	// selected with [CTRL]+[SHIFT]+[C].
	switch ( treeSelectedType )
	{
//		case otElement :
//		{
//			// Code for when an element is selected
//			var theElement as EA.Element;
//			theElement = Repository.GetTreeSelectedObject();
//						
//			break;
//		}
		case otPackage :
		{
			// Code for when a package is selected
			var thePackage as EA.Package;
			thePackage = Repository.GetTreeSelectedObject();
			selectedPackage = thePackage; // To be able to use it in callbacks
			ImportDWH(thePackage);
			break;
		}
//		case otDiagram :
//		{
//			// Code for when a diagram is selected
//			var theDiagram as EA.Diagram;
//			theDiagram = Repository.GetTreeSelectedObject();
//			
//			break;
//		}
//		case otAttribute :
//		{
//			// Code for when an attribute is selected
//			var theAttribute as EA.Attribute;
//			theAttribute = Repository.GetTreeSelectedObject();
//			
//			break;
//		}
//		case otMethod :
//		{
//			// Code for when a method is selected
//			var theMethod as EA.Method;
//			theMethod = Repository.GetTreeSelectedObject();
//			
//			break;
//		}
		default:
		{
			// Error message
			Session.Prompt( "This script does not support items of this type.", promptOK );
		}
	}
}

function ImportDWH(thePackage /*: EAElement */) {
	Repository.EnsureOutputVisible( "Script" );
	LOGInfo("ImportDWH **:"+thePackage.Name);
    /*
    var x as EA.Package;
	
    x = Repository.GetElementByGuid("{A38E1ED1-2517-4040-9AB7-2A61EA8F9D68}");	
	LOGElement(x);
	*/
	/*
	LOGElement(Repository.GetTreeSelectedObject().Packages.GetAt(0));
	LOGElement(Repository.GetTreeSelectedObject().Packages.GetAt(1));
	
    LOGInfo(CountElements(Repository.GetTreeSelectedObject().Packages.GetAt(0)));
	LOGInfo(CountElements(Repository.GetTreeSelectedObject().Packages.GetAt(1)));
	LOGInfo("Done");
	//return;
	*/
	CSV_DELIMITER = ";";
	
	if (thePackage.Name == "JOBS") {
	  // Select the file with the elements
	  var elementsFilename = DLGOpenFile( "All Files |*.*|CSV|*.csv", 2 );
      LOGInfo("ElementsFilename:"+elementsFilename);
      // Import Elements
	  execMode = 0; //"JOBS"
	  LOGInfo("Importing JOBS");
	  Repository.EnableUIUpdates = false;
	  CSVIImportFile( elementsFilename, true);
	  Repository.RefreshModelView(0);
	}
	
    if (thePackage.Name == "DATASETS") {
	  // Select the file with the elements
	  var elementsFilename = DLGOpenFile( "All Files |*.*|CSV|*.csv", 2 );
      LOGInfo("ElementsFilename:"+elementsFilename);
      // Import Elements
	  execMode = 1; //"JOBS"
	  LOGInfo("Importing DATASETS");
	  Repository.EnableUIUpdates = false;
	  CSVIImportFile( elementsFilename, true);
	  Repository.RefreshModelView(0);
	}
	
    if (thePackage.Name =="Import Area") {	
      // Select the file with the relationships
	  var relationshipsFilename = DLGOpenFile( "All Files |*.*|CSV|*.csv", 2 );
      LOGInfo("RelationshipsFilename:"+relationshipsFilename);
	
	  // Import Relationships
	  execMode = 2; // "RELATIONSHIPS"
	  LOGInfo("Importing Relationships");
	  Repository.EnableUIUpdates = false;
	  CSVIImportFile( relationshipsFilename, true);
	  Repository.RefreshModelView(0);
	}
	
	LOGError("Unable to import anything");
}

/* CSV Import can be performed by calling the function CSVIImportFile(). CSVIImportFile() requires
 * that the function OnRowImported() be defined in the user's script to be used as a callback
 * whenever row data is read from the CSV file. The user defined OnRowImported() can query for 
 * information about the current row through the functions CSVIContainsColumn(), 
 * CSVIGetColumnValueByName() and CSVIGetColumnValueByNumber().
 */

/* Callback */
function OnRowImported() {
	LOGInfo("OnRowImported");
	switch (execMode) {
		case 0:
			ImportOneJob();	
			break;
		case 1:
			ImportOneDataset();	
			break;
		case 2:
            ImportOneRelationship();
			break;
		default:
	}
}

function ImportOneJob() {
	// Create new element
    // Set its values
	LOGInfo("ImportOneElement");
	var name = CSVIGetColumnValueByName("JOB");
	LOGInfo("JOB="+name);
	var theElement as EA.Element;
	
	selectedPackage.Elements.AddNew(name,"Class");
	
}

function LOGElement(theElement) {
	LOGInfo("Element:"+theElement.Name+" ("+theElement.Type+")");
}

function ImportOneDataset() {
	// Create new element
    // Set its values
	LOGInfo("ImportOneElement");
	var name = CSVIGetColumnValueByName("DATASET");
	LOGInfo("DATASET="+name);
	var theElement as EA.Element;
	
	selectedPackage.Elements.AddNew(name,"Class");
}

function findElementByName(thePackage, name) {
	
	var elem as EA.Element
	LOGInfo("findElementByName("+name+")");
	LOGElement(thePackage);

	LOGInfo("About to search");
	
	elem = thePackage.Elements.GetByName(name);
	
	if (elem == null) {
		LOGInfo("Not found");
		return -1;
	} else {
		LOGInfo("Found:"+elem.ElementID);
		return elem.ElementID;
	}
	return -2;
}

function findJobByName(name) {
   LOGInfo("FindJobByName("+name+")");
//   LOGElement(thePackage);
   //return findElementByName(Repository.GetElementByGuid("{CDE16555-7D9D-4593-BDD6-339B7D51A08F}"),name);
   return findElementByName(Repository.GetTreeSelectedObject().Packages.GetAt(0),name);
}
   
function findDatasetByName(name) {
   LOGInfo("FindDatasetByName("+name+")");
  // var thePackage = Repository.GetElementByGuid("{A38E1ED1-2517-4040-9AB7-2A61EA8F9D68}");
  // LOGElement(thePackage);
  //return findElementByName(Repository.GetElementByGuid("{A38E1ED1-2517-4040-9AB7-2A61EA8F9D68}"),name);
  
  return findElementByName(Repository.GetTreeSelectedObject().Packages.GetAt(1),name);

}

function ImportOneRelationship() {
	// Find source element by name
	// Find target element by name
	// Create a relationship between the two
	LOGInfo("ImportOneRelationship");
	
	var source = CSVIGetColumnValueByName("JOB");
	var destination = CSVIGetColumnValueByName("DATASET");
	
	LOGInfo("JOB="+source+" - DATASET="+destination);

	var con As EA.Connector;
	

	LOGInfo("##### "+findJobByName(source));
	
	var srcID = findJobByName(source);
	var trgID = findDatasetByName(destination);
	
	LOGInfo("src:"+srcID+" - trg:"+trgID);

	con = Repository.GetElementByID(srcID).Connectors.AddNew ("", "Association");
    con.SupplierID = trgID;

   if (!con.Update) {
       LOGError("Connector: ("+source+"->"+destination+")"+con.GetLastError);
   }
   
}

function CountElements( thePackage )
{
	var count = 0;
	
	// Cast thePackage to EA.Package so we get intellisense
	var contextPackage as EA.Package;
	contextPackage = thePackage;
	
	// Iterate through all child packages
	var childPackageEnumerator = new Enumerator( contextPackage.Packages );
	while ( !childPackageEnumerator.atEnd() )
	{
		var currentPackage as EA.Package;
		currentPackage = childPackageEnumerator.item();
		
		// Recursively process child packages
		count = count + CountElements( currentPackage );
		
		childPackageEnumerator.moveNext();
	}
	
	// Add this package's element count to the counter
	return count + contextPackage.Elements.Count;
}


OnProjectBrowserScript();

