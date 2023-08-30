using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITrainingPlotExport
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            var sheets=new FilteredElementCollector(doc)
                            .WhereElementIsNotElementType()
                            .OfClass(typeof(ViewSheet))
                            .Cast<ViewSheet>()
                            .ToList();
            var groupedSheets = sheets.GroupBy(sheet => doc.GetElement(new FilteredElementCollector(doc, sheet.Id)
                                                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                                                    .FirstElementId()).Name);
            var viweSetes = new List<ViewSet>();

            PrintManager printManager = doc.PrintManager;


            printManager.PrintRange = PrintRange.Select;


            ViewSheetSetting viewSheetSetting=printManager.ViewSheetSetting;

            foreach (var groupedSheet in groupedSheets)
            {
                if(groupedSheet.Key==null) continue;

                var viewSet=new ViewSet();

                var sheetsOfGroup=groupedSheet.Select(s=>s).ToList();
                foreach (var sheet in sheetsOfGroup)
                {
                    viewSet.Insert(sheet);
                }

                viweSetes.Add(viewSet);

                printManager.PrintRange = PrintRange.Select;
                viewSheetSetting.CurrentViewSheetSet.Views = viewSet;

                using(var ts=new Transaction(doc,"Create view set"))
                {
                    ts.Start();
                    viewSheetSetting.SaveAs($"{groupedSheet.Key}_{Guid.NewGuid()}");
                    ts.Commit();
                }
            }


            return Result.Succeeded;
        }
    }
}
