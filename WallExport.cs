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
            IList<Element> _elements = collector.OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements();
            
            // I create this class to structure the data we want to export
            var csv = new StringBuilder();
            string headers = string.Format("{0},{1},{2}", "ElementName", "Category", "Id");
            csv.AppendLine(headers);

            foreach (Element elem in _elements)
            {
                Parameter param = elem.LookupParameter("Comments");
                if ( param != null && param.AsString() != "")
                {
                    string newLine = string.Format("{0},{1},{2}", elem.Name, elem.Category.Name, elem.Id.ToString());
                    csv.AppendLine(newLine);
                }
            }
            File.WriteAllText(@"C:/fran.csv", csv.ToString());


            return Result.Succeeded;

        }
    }
}
