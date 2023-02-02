using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;

namespace RevitAPITraining_DataOutputXLSX
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            //Сбор всех труб
            List<Pipe> pipes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .WhereElementIsNotElementType()
                .Cast<Pipe>()
                .ToList();

            //Указание пути, названия файла
            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "pipes.xlsx");

            //Создание стрим(запись данных в файл)
            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                //Создание файла ексель
                IWorkbook workbook = new XSSFWorkbook();

                //Создание листа
                ISheet sheet = workbook.CreateSheet("Лист1");

                int rowIndex = 0;
                foreach (var pipe in pipes)
                {
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipe.PipeType.Name);
                    sheet.SetCellValue(rowIndex, columnIndex: 1, pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsValueString());
                    sheet.SetCellValue(rowIndex, columnIndex: 2, pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsValueString());
                    sheet.SetCellValue(rowIndex, columnIndex: 3, pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString());
                    rowIndex++;
                }
                //Запись данных в файл
                workbook.Write(stream);

                //Закрытие файла
                workbook.Close();
            }

            //Открытие файла в эксель
            System.Diagnostics.Process.Start(excelPath);

            return Result.Succeeded;
        }
    }
}
