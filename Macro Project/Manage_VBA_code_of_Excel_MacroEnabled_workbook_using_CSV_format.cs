using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCsvDemo
{
    public class VbaCsvManager
    {
        public static void Run()
        {
            // Determine CSV path relative to the executable directory
            string csvPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "VbaModules.csv");

            // If the CSV does not exist, create a sample one
            if (!File.Exists(csvPath))
            {
                using (var writer = new StreamWriter(csvPath))
                {
                    // ModuleName,Code (code uses escaped line breaks)
                    writer.WriteLine("SampleModule,Sub Test()\\r\\n    MsgBox \"Hello from VBA\"\\r\\nEnd Sub");
                }
            }

            // Load CSV workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook csvWorkbook = new Workbook(csvPath, loadOptions);

            // Create a new workbook that will hold VBA project
            Workbook macroWorkbook = new Workbook();

            // Ensure the workbook has a VBA project by saving as .xlsm and reloading
            if (macroWorkbook.VbaProject == null)
            {
                string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsm");
                macroWorkbook.Save(tempPath, SaveFormat.Xlsm);
                macroWorkbook = new Workbook(tempPath);
                File.Delete(tempPath);
            }

            Worksheet csvSheet = csvWorkbook.Worksheets[0];
            int maxRow = csvSheet.Cells.MaxDataRow;

            for (int row = 0; row <= maxRow; row++)
            {
                string moduleName = csvSheet.Cells[row, 0].StringValue?.Trim();
                if (string.IsNullOrEmpty(moduleName))
                    continue;

                string rawCode = csvSheet.Cells[row, 1].StringValue ?? string.Empty;
                string vbaCode = rawCode.Replace("\\r\\n", "\r\n")
                                        .Replace("\\n", "\n")
                                        .Replace("\\r", "\r");

                int moduleIndex = macroWorkbook.VbaProject.Modules.Add(VbaModuleType.Procedural, moduleName);
                VbaModule vbaModule = macroWorkbook.VbaProject.Modules[moduleIndex];
                vbaModule.Codes = vbaCode;
            }

            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WorkbookWithVba.xlsm");
            macroWorkbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"VBA modules imported from CSV and saved to '{outputPath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VbaCsvManager.Run();
        }
    }
}