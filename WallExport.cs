using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using System.IO;
using CsvHelper;
using System.Globalization;

namespace ExportToExcel
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class WallExport : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            Document doc = commandData.Application.ActiveUIDocument.Document;

            FilteredElementCollector collector = new FilteredElementCollector(doc);


            List<Parameter> ElemWidthList = new List<Parameter>();
           
            foreach (Element e in AllWalls)
            {
                
                ElemWidthList.Add(e.LookupParameter("Width"));
                

                    if (parameterValue != "" || parameterValue != "")
                    {
                        filteredElements.Add(element); // you have all elements that need to fill this parameter.
                        filteredElementIds.Add(element.Id); // or you could save the id directly.

                        // or write a default value 
                        // elementParameter.Set("default value");                    

                    }
                }
            }









            //using (var csvWriter = new CsvHelper.CsvWriter(writer)
            //{
            //    foreach (string s in AllWallsExcelKnows)
            //    {
            //        csvWriter.WriteField(s);

            //    }

            //}






            return Result.Succeeded;

        }
    }

}
