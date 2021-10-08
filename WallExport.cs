using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;

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


            List<Parameter> ElemWidthList = new List<Parameter>();
           
            foreach (Element e in AllWalls)
            {
                
                ElemWidthList.Add(e.LookupParameter("Width"));
                

            }



            return Result.Succeeded;

        }
    }

}
