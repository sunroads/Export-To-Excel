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

            // Replace this value to iterate on other categories.
            BuiltInCategory builtInCategory = BuiltInCategory.OST_Walls;

            ElementCategoryFilter elementCategoryFilter = new ElementCategoryFilter(builtInCategory);
            
            IList <Element> allElements = collector.WherePasses(elementCategoryFilter).WhereElementIsNotElementType().ToElements();

            // Create an empty list to save all the elements that don´t have a value on the parameter width
            List<Element> filteredElements = new List<Element>();
            List<ElementId> filteredElementIds = new List<ElementId>();


            foreach (Element element in filteredElements)
            {
                string parameterName = "Width";
                Parameter elementParameter = element.LookupParameter(parameterName);

                // with this if, you check if the parameter exists on this element.
                if (elementParameter != null)
                {
                    //Then you read the value
                    string parameterValue = elementParameter.AsValueString();

                // sometimes the value is not an string, so take a look on this posibilities, you may need to add a switch to identify the storage type before read it.

                // https://www.revitapidocs.com/2022/0b04b80d-b318-986e-48cc-835d0dda76e5.htm

                    if (parameterValue != "" || parameterValue != "")
                    {
                        filteredElements.Add(element); // you have all elements that need to fill this parameter.
                        filteredElementIds.Add(element.Id); // or you could save the id directly.

                        // or write a default value 
                        // elementParameter.Set("default value");                    

                    }
                }
            }

            return Result.Succeeded;

        }
    }

}
