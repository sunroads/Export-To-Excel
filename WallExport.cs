using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System.Globalization;


namespace ExportToExcel
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class WallExport : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Document doc = commandData.Application.ActiveUIDocument.Document;

            //COLLECTOR
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            IList<Element> _elementsDB = collector.OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements();
            ICollection<ElementId> _elementsIds = collector.OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElementIds();


            //foreach (Element e in _elementsDB)
            //{

            //    List<Parameter> _NoHeightParameter = new List<Parameter>();

            //    if (e.LookupParameter("Неприсоединенная высота") == null)
            //    {

            //        TaskDialog.Show("No Height", e.Name + " " + e.Id + " doesn't have Height Parameter");
            //        //_NoHeightParameter.Add(e.LookupParameter("Неприсоединенная высота"));
            //        //_NoHeightParameter.Count
            //    }                               

            //}   


            List<double> _pHeight = new List<double>();


            foreach (Element e in _elementsDB)
            {


                if (e.LookupParameter("Неприсоединенная высота") != null)
                {
                    _pHeight.Add(UnitUtils.ConvertFromInternalUnits(e.LookupParameter("Неприсоединенная высота").AsDouble(), DisplayUnitType.DUT_MILLIMETERS));
                }
                else
                {
                    continue;
                }

            }


            //    ////////////////////////////////////////////////////////////////EXCEL

            ExcelPackage package = new ExcelPackage();
            ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet");
            var file = new FileInfo(@"D:\WallsLE 2 ячейка.xlsx");

            using (ExcelRange Rng = ws.Cells["A1:E2"])
            {
                //Indirectly access ExcelTableCollection class
                ExcelTable table = ws.Tables.Add(Rng, "Table");


                table.Columns[0].Name = "Name";
                table.Columns[1].Name = "ID";
                table.Columns[2].Name = "Height";

            }

            //1-st row
            //int col = 1;
            //for (int row = 1; row <= _elementsDB.Count; row++)
            //{
            //    ws.Cells[row, col].Value = _elementsDB[row-1].Name;
            //}
            //2-nd row
            int col = 1;
            for (int row = 2; row - 1 <= _elementsDB.Count; row++)
            {
                ws.Cells[row, col].Value = _elementsDB[row - 2].Name;
            }

            int col2 = 2;
            for (int row = 2; row - 1 <= _elementsIds.Count; row++)
            {
                ws.Cells[row, col2].Value = _elementsIds.ElementAt(row - 2);
            }

            int col3 = 3;
            for (int row = 2; row - 1 <= _elementsIds.Count; row++)
            {
                ws.Cells[row, col3].Value = _pHeight[row - 2];
            }

            package.SaveAs(new FileInfo(file.ToString()));

            return Result.Succeeded;

        }
    }
}
