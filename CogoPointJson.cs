using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.Civil.DatabaseServices;
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

namespace civil3dCogoPoints
{
    class CogoPointJson
    {
        public static Dictionary<string, object> geolocationCapture( CogoPointCollection passedCogoCollection)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var acDB = doc.Database;

            // Creat a json file of the Corner Record Points found in the CogoPointCollection
            // Retrieves the Long/Lat of the Point and converts to decimal degrees
            // Confirms that the Name field is filled correctly (cr x)
            
            using (var trans = acDB.TransactionManager.StartTransaction())
            {
                Dictionary<string, object> cogoPointJson = new Dictionary<string, object>();

                foreach (ObjectId cogoPointRecord in passedCogoCollection)
                {
                    CogoPoint cogoPointItem = trans.GetObject(cogoPointRecord, OpenMode.ForRead) as CogoPoint;

                    Match cogoMatch = Regex.Match(cogoPointItem.PointName, "^(\\s*cr\\s*\\d\\d*)$",
                        RegexOptions.IgnoreCase);

                    if (cogoMatch.Success)
                    {
                        Dictionary<String, object> cogoPointGeolocation = new Dictionary<string, object>();

                        //convert the Lat/Long from Radians to Decimal Degrees
                        double rad2DegLong = (cogoPointItem.Longitude * 180) / Math.PI;
                        double rad2DegLat = (cogoPointItem.Latitude * 180) / Math.PI;

                        cogoPointGeolocation.Add("Corner_Type_c", "Other");
                        cogoPointGeolocation.Add("Geolocation_Longitude_s", rad2DegLong);
                        cogoPointGeolocation.Add("Geolocation_Latitude_s", rad2DegLat);
                        cogoPointGeolocation.Add("Full Description", cogoPointItem.FullDescription);

                        cogoPointJson.Add(cogoPointItem.PointName.Trim().ToString().ToLower().Replace(" ", ""),
                            cogoPointGeolocation);
                    }
                }

                using (var writer = File.CreateText("CogoPointsGelocation.json"))
                {
                    string strResultJson = JsonConvert.SerializeObject(cogoPointJson,
                        Formatting.Indented);
                    writer.WriteLine(strResultJson);
                }
                trans.Commit();

                return cogoPointJson;
            }
        }
    }
}
