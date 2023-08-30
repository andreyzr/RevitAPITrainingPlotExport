using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAPITrainingExportToImage
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            using (var ts = new Transaction(doc, "export img"))
            {
                ts.Start();

                string desktop_path = Environment.GetFolderPath(
                Environment.SpecialFolder.Desktop);

                ViewPlan viewPlan = new FilteredElementCollector(doc)
                                    .OfClass(typeof(ViewPlan))
                                    .Cast<ViewPlan>()
                                    .FirstOrDefault(v => v.ViewType == ViewType.FloorPlan &&
                                    v.Name.Equals("Level 1"));

                string filepath = Path.Combine(desktop_path,
                        viewPlan.Name);

                ImageExportOptions imgOptions = new ImageExportOptions();

                List<ElementId> ImageExportList = new List<ElementId>();

                ImageExportList.Add(viewPlan.Id);

                imgOptions.SetViewsAndSheets(ImageExportList);

                imgOptions.HLRandWFViewsFileType = ImageFileType.PNG;
                imgOptions.FilePath = filepath;
                imgOptions.ShadowViewsFileType = ImageFileType.PNG;

                doc.ExportImage(imgOptions);
                ts.Commit();


            }


            return Result.Succeeded;
        }

    }
}
