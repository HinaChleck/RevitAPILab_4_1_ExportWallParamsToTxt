using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace RevitAPILab_4_1_ExportWallParamsToTxt
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        //Экспортирует Марку и объемы стен в txt
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            #region Диалог сохранения файла
            var saveFileDialog = new SaveFileDialog
            {
               OverwritePrompt = true,
               InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
               Filter = "All files (*.*)|*.*",
               FileName =   "wallInfo.txt",
               DefaultExt = ".txt"
            };

            string selectedFilePath = string.Empty;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = saveFileDialog.FileName;
            }

            if (string.IsNullOrEmpty(selectedFilePath))
                return Result.Cancelled;

            #endregion

            List<Wall> allWalls = new FilteredElementCollector(doc)
                 .OfClass(typeof(Wall))
                 .Cast<Wall>()
                 .ToList();
                                                  
            string allText = string.Empty;
            foreach (var wall in allWalls)
            {
           
                string wallType = wall.LookupParameter("Марка").AsString();
                string WallVolume = UnitUtils.ConvertFromInternalUnits(wall.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble(), DisplayUnitType.DUT_CUBIC_METERS).ToString();
                allText += $"{wallType}\t{WallVolume}{Environment.NewLine}";
                              
            }

            File.WriteAllText(selectedFilePath, allText);
            TaskDialog.Show("Экспорт", $"Файл сохранен по указанному пути:\n{selectedFilePath}\n{allText}");

            return Result.Succeeded;
        }

    }
}
