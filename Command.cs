using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil;
using Autodesk.Civil.ApplicationServices;
using CivilDB = Autodesk.Civil.DatabaseServices;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using civil3dCogoPoints;
using autoCadDbText;

[assembly: CommandClass(typeof(CrxApp.Commands))]
[assembly: ExtensionApplication(null)]
namespace CrxApp
{
    #region Commands
    public class Commands
    {
        [CommandMethod("MyCommands", "OCPWCR", CommandFlags.Modal)]
        static public void CornerRecordData()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                var acDB = doc.Database;

                using (var trans = acDB.TransactionManager.StartTransaction())
                {
                    DBDictionary layoutPages = (DBDictionary)trans.GetObject(acDB.LayoutDictionaryId, 
                        OpenMode.ForRead);
  
                    // Handle Corner Record meta data dictionary extracted from Properties and Content
                    Dictionary<String, object> cornerRecordForms = new Dictionary<string, object>();

                    CivilDB.CogoPointCollection cogoPointsColl = CivilDB.CogoPointCollection.GetCogoPoints(doc.Database);
                    var cogoPointCollected = CogoPointJson.geolocationCapture(cogoPointsColl);

                    List<string> layoutNamesList = new List<string>();

                    foreach (DBDictionaryEntry layoutPage in layoutPages)
                    {
                        var crFormItems = layoutPage.Value.GetObject(OpenMode.ForRead) as Layout;
                        var isModelSpace = crFormItems.ModelType;

                        ObjectIdCollection textObjCollection = new ObjectIdCollection();

                        // Formatted Dictionary to create JSON output
                        Dictionary<string, object> textObjResults = new Dictionary<string, object>();

                        if (isModelSpace != true)
                        {
                            BlockTableRecord blkTblRec = trans.GetObject(crFormItems.BlockTableRecordId,
                                OpenMode.ForRead) as BlockTableRecord;

                            layoutNamesList.Add(crFormItems.LayoutName.Trim().ToString().ToLower().Replace(" ", ""));

                            foreach (ObjectId btrId in blkTblRec)
                            {
                                if (btrId.ObjectClass.DxfName.Contains("TEXT") || btrId.ObjectClass.DxfName.Contains("MTEXT"))
                                {
                                    textObjCollection.Add(btrId);
                                }
                            }

                            // List of the Dictionary txtProps
                            List<object> crFormElements = new List<object>();
                            List<object> crFormInputs = new List<object>();
                            List<ObjectId> crFormInputId = new List<ObjectId>(); 

                            foreach (ObjectId txtItem in textObjCollection)
                            {
                                // Dictionary to collect Contents of the Corner Records Sheet 
                                Dictionary<string, object> txtProps = new Dictionary<string, object>();

                                var txtItemName = trans.GetObject(txtItem, OpenMode.ForRead) as Entity;

                                if (txtItemName.Layer == "-SHEET")
                                {
                                    if (txtItem.ObjectClass.DxfName == "MTEXT")
                                    {
                                        var mtextValues = trans.GetObject(txtItem, OpenMode.ForRead)
                                            as MText;

                                        //txtProps.Add("Class Type", "MTEXT");
                                        //txtProps.Add("Layer Name", mtextValues.Layer);
                                        txtProps.Add("Content", mtextValues.Text);
                                        //txtProps.Add("X", mtextValues.Location.X);
                                        //txtProps.Add("Y", mtextValues.Location.Y);

                                        crFormElements.Add(txtProps);
                                    }
                                    else if (txtItem.ObjectClass.DxfName == "TEXT")
                                    {
                                        //Capture Brief Legal Description here
                                        
                                        var textValues = trans.GetObject(txtItem, OpenMode.ForRead)
                                            as DBText;

                                        //txtProps.Add("Class Type", "TEXT");
                                        //txtProps.Add("Layer Name", textValues.Layer);
                                        txtProps.Add("Content", textValues.TextString);
                                        //txtProps.Add("X", textValues.Position.X);
                                        //txtProps.Add("Y", textValues.Position.Y);

                                        crFormElements.Add(txtProps);
                                    }
                                }
                                else if (txtItemName.Layer == "$--SHT-ANNO")
                                {
                                    if (txtItem.ObjectClass.DxfName == "MTEXT")
                                    {
                                        var mtextValues = trans.GetObject(txtItem, OpenMode.ForRead)
                                            as MText;

                                        //txtProps.Add("Class Type", "MTEXT");
                                        //txtProps.Add("Layer Name", mtextValues.Layer);
                                        txtProps.Add("Content", mtextValues.Text);
                                        //txtProps.Add("X", mtextValues.Location.X);
                                        //txtProps.Add("Y", mtextValues.Location.Y);

                                        crFormInputs.Add(txtProps);
                                        crFormInputId.Add(txtItem);
                                    }
                                    else if (txtItem.ObjectClass.DxfName == "TEXT")
                                    {
                                        var textValues = trans.GetObject(txtItem, OpenMode.ForRead)
                                            as DBText;

                                        //txtProps.Add("Class Type", "TEXT");
                                        //txtProps.Add("Layer Name", textValues.Layer);
                                        txtProps.Add("Content", textValues.TextString);
                                        //txtProps.Add("X", textValues.Position.X);
                                        //txtProps.Add("Y", textValues.Position.Y);

                                        crFormInputs.Add(txtProps);
                                        crFormInputId.Add(txtItem);
                                    }
                                }
                            }
                            var briefLegalCollected = BriefLegalDescription.briefLegalCapture(textObjCollection, crFormInputId);
                            textObjResults.Add("Legal_Description_c", briefLegalCollected);
                            //textObjResults.Add("Form Inputs", crFormInputs);
                            //textObjResults.Add("Form Elements", crFormElements);
                            cornerRecordForms.Add(crFormItems.LayoutName.Trim().ToString().ToLower().Replace(" ",""), textObjResults);
                        }
                    }

                    // Checks to see whether the points from the cogo point collection exist within 
                    // the layout by searching for the correct collection key and layout name
                    List<string> cogoPointCollectedCheck = cogoPointCollected.Keys.ToList();
                    List<bool> boolCheckResults = new List<bool>();

                    IEnumerable<string> cogoPointNameCheck = layoutNamesList.Except(cogoPointCollectedCheck);
                    List<string> cogoPointNameCheckResults = cogoPointNameCheck.ToList();
                    var layoutNameChecker = new Regex("^(\\s*cr\\s*\\d\\d*)$");

                    if (!cogoPointNameCheckResults.Where(f => layoutNameChecker.IsMatch(f)).ToList().Any())
                    {
                        boolCheckResults.Add(true);
                    }
                    else
                    {
                        foreach (string cogoPointNameResultItem in cogoPointNameCheckResults)
                        {
                            Match layoutNameMatch = Regex.Match(cogoPointNameResultItem, "^(\\s*cr\\s*\\d\\d*)$",
                                RegexOptions.IgnoreCase);

                            if (layoutNameMatch.Success)
                            {
                                string layoutNameX = layoutNameMatch.Value;
                                ed.WriteMessage("\nLayout Named {0} does not have an associated cogo point", layoutNameX);
                            }
                        }
                        boolCheckResults.Add(false);
                    }


                    IEnumerable<string> layoutNameCheck = cogoPointCollectedCheck.Except(layoutNamesList);
                    List<string> layoutNameCheckResults = layoutNameCheck.ToList();
                    var cogoNameChecker = new Regex("^(\\s*cr\\s*\\d\\d*)$");
                    
                    
                    // If the layout name has any value other than CR == PASS
                    // If CR point exists and does not match then throw an error for user to fix
                    if(!layoutNameCheckResults.Where(f => cogoNameChecker.IsMatch(f)).ToList().Any())
                    {
                        boolCheckResults.Add(true);
                    }
                    else // Found a CR point that DID NOT match a layout name 
                    {
                        foreach (string layoutNameCheckResultItem in layoutNameCheckResults)
                        {
                            Match cogoNameMatch = Regex.Match(layoutNameCheckResultItem, "^(\\s*cr\\s*\\d\\d*)$",
                                RegexOptions.IgnoreCase);

                            if (cogoNameMatch.Success)
                            {
                                string cogoNameX = cogoNameMatch.Value;
                                ed.WriteMessage("\nCorner Record point named {0} does not have an associated Layout", 
                                    cogoNameX);
                            }
                        }
                        boolCheckResults.Add(false);
                    }

                    // Output JSON file to BIN folder
                    // IF there are two true booleans in the list then add the data to the corresponding keys (cr1 => cr1)
                    if ((boolCheckResults.Count(v => v == true)) == 2)
                    {
                        foreach (string cornerRecordFormKey in cornerRecordForms.Keys)
                        {
                            if (cogoPointCollected.ContainsKey(cornerRecordFormKey))
                            {
                                //ed.WriteMessage("THIS SHIT FINALLY WORKS FOR {0}", cornerRecordFormKey);
                                var cogoFinal = (Dictionary<string, object>)cornerRecordForms[cornerRecordFormKey];
                                var cogoFinalType = ((Dictionary<string, object>)cogoPointCollected[cornerRecordFormKey])
                                    ["Corner_Type_c"];
                                var cogoFinalLong = ((Dictionary<string, object>)cogoPointCollected[cornerRecordFormKey])
                                    ["Geolocation_Longitude_s"];
                                var cogoFinalLat = ((Dictionary<string, object>)cogoPointCollected[cornerRecordFormKey])
                                    ["Geolocation_Latitude_s"];

                                cogoFinal.Add("Corner_Type_c", cogoFinalType);
                                cogoFinal.Add("Geolocation_Longitude_s", cogoFinalLong);
                                cogoFinal.Add("Geolocation_Latitude_s", cogoFinalLat); 
                            }
                        }   
                        using (var writer = File.CreateText("CornerRecordForms.json"))
                        {
                            string strResultJson = JsonConvert.SerializeObject(cornerRecordForms,
                                Formatting.Indented);
                            writer.WriteLine(strResultJson);
                        }
                    }
                    trans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("Error: {0}", ex);
            }
        }
    }
    #endregion
}