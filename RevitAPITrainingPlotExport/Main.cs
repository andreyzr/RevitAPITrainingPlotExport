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

            using(var ts =new Transaction(doc,"export dwg"))
            {
                ts.Start();
                ViewPlan viewPlan = new FilteredElementCollector(doc)
                                    .OfClass(typeof(ViewPlan))
                                    .Cast<ViewPlan>()
                                    .FirstOrDefault( v => v.ViewType == ViewType.FloorPlan &&
                                    v.Name.Equals("Level 1"));
                var dwgOption = new DWGExportOptions();
                doc.Export(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "export.dwg",
                    new List<ElementId> { viewPlan.Id },dwgOption);
                ts.Commit();
            }

           
            return Result.Succeeded;
        }

        public void BatchPrint(Document doc)
        {
            var sheets = new FilteredElementCollector(doc)
                            .WhereElementIsNotElementType()
                            .OfClass(typeof(ViewSheet))
                            .Cast<ViewSheet>()
                            .ToList();
            var groupedSheets = sheets.GroupBy(sheet => doc.GetElement(new FilteredElementCollector(doc, sheet.Id)
                                                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                                                    .FirstElementId()).Name);
            var viweSetes = new List<ViewSet>();

            PrintManager printManager = doc.PrintManager;
            printManager.SelectNewPrintDriver("PDF24");
            printManager.PrintRange = PrintRange.Select;
            ViewSheetSetting viewSheetSetting = printManager.ViewSheetSetting;

            foreach (var groupedSheet in groupedSheets)
            {
                if (groupedSheet.Key == null) continue;//

                var viewSet = new ViewSet();

                var sheetsOfGroup = groupedSheet.Select(s => s).ToList();
                foreach (var sheet in sheetsOfGroup)
                {
                    viewSet.Insert(sheet);
                }

                viweSetes.Add(viewSet);

                printManager.PrintRange = PrintRange.Select;
                viewSheetSetting.CurrentViewSheetSet.Views = viewSet;

                using (var ts = new Transaction(doc, "Create view set"))
                {
                    ts.Start();
                    viewSheetSetting.SaveAs($"{groupedSheet.Key}_{Guid.NewGuid()}");//Создание набора листов
                    ts.Commit();
                }


                bool isFormatSelected = false;
                foreach (PaperSize paperSize in printManager.PaperSizes)//Вывод на печать
                {

                    if (string.Equals(groupedSheet.Key, "А4К") &&
                        string.Equals(paperSize.Name, "A3"))
                    {
                        printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PaperSize = paperSize;
                        printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PageOrientation = PageOrientationType.Portrait;
                        isFormatSelected = true;
                    }
                    else if (string.Equals(groupedSheet.Key, "А3А") &&
                        string.Equals(paperSize.Name, "A3"))
                    {
                        printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PaperSize = paperSize;
                        printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PageOrientation = PageOrientationType.Landscape;
                        isFormatSelected = true;
                    }
                }
                if (!isFormatSelected)
                {
                    TaskDialog.Show("Ошибка", "Не найден формат");
                }
                printManager.CombinedFile = false;// сохранение в разные файлы
                printManager.SubmitPrint();//Печатаем с текущими настройками
            }
        }   //Печать и  созданеи наборов 
    }
}
