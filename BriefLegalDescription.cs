using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;

using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil;
using Autodesk.Civil.ApplicationServices;
using CivilDB = Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

namespace autoCadDbText
{
    class BriefLegalDescription
    {
        public static string briefLegalCapture(ObjectIdCollection passedFormElements, 
            List<ObjectId> passedCrFormInputs)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var acDB = doc.Database;

            List<string> briefLegalReturn = new List<string>();

            using (var trans = acDB.TransactionManager.StartTransaction())
            {
                foreach (ObjectId FormElement in passedFormElements)
                {
                    if (FormElement.ObjectClass.DxfName == "TEXT")
                    {
                        var FormElementValue = trans.GetObject(FormElement, OpenMode.ForRead)
                        as DBText;
                    
                        if (FormElementValue.Layer == "-SHEET")
                        {
                            var textValues = (DBText)trans.GetObject(FormElement, OpenMode.ForRead);

                            if (textValues.TextString == "Brief Legal Description")
                            {
                                foreach (ObjectId inputItem in passedCrFormInputs)
                                {
                                    //var inputItemValue = trans.GetObject(inputItem, OpenMode.ForRead) as Entity;
                                    if (inputItem.ObjectClass.DxfName == "TEXT")
                                    {
                                        var inputItemText = trans.GetObject(inputItem, OpenMode.ForRead) as DBText;
                                        double crElementX = Math.Abs(textValues.Position.X - inputItemText.Position.X);
                                        double crElementY = Math.Abs(textValues.Position.Y - inputItemText.Position.Y);

                                        if ((crElementX < 5) && (crElementY < 0.13) && 
                                            (textValues.Position.X < inputItemText.Position.X))
                                        {
                                            briefLegalReturn.Add(inputItemText.TextString);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return briefLegalReturn[0];
        }    
    }
}
