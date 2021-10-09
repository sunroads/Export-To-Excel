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
            ElementCategoryFilter FilterWalls = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            IList <Element> AllWalls = collector.WherePasses(FilterWalls).WhereElementIsNotElementType().ToElements();

            
            List<string> AllWallsExcelKnows = new List<string>();

            foreach (Element e in AllWalls)
            {

                AllWallsExcelKnows.Add(e.Name.ToString());
                
            }

            //List<Parameter> ElemWidthList = new List<Parameter>();

            //foreach (Element e in AllWalls)
            //{

            //    ElemWidthList.Add(e.LookupParameter("Width"));


            //}

            using (var writer = new StreamWriter("C:\\Users\\b-filippov\\Desktop\\test\\Excel\\elements.csv"))


            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                foreach (var s in AllWallsExcelKnows)
                {
                    csv.WriteField(s);
                }

                //writer.Flush();
                

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
