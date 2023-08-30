using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPIExportToNWC
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;


                View viewPlan = new FilteredElementCollector(doc)
                          .OfClass(typeof(View))
                          .Cast<View>()
                          .FirstOrDefault(v =>v.Name.Equals("{3D}"));

                var nwcOption = new NavisworksExportOptions();

                nwcOption.ExportScope = NavisworksExportScope.View;
                nwcOption.ViewId = viewPlan.Id;

                doc.Export(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "export.nwc",
                     nwcOption);



            return Result.Succeeded;
        }

    }
}
