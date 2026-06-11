using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (default format is Xlsx)
                Workbook workbook = new Workbook();

                // Access the first worksheet and give it a friendly name
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Put some sample data (optional, just to have cells to format)
                sheet.Cells["A1"].PutValue("Sample");
                sheet.Cells["A2"].PutValue("Data");

                // Get the existing ThisWorkbook module or add it if missing
                VbaModule vbaModule = null;
                foreach (VbaModule mod in workbook.VbaProject.Modules)
                {
                    if (mod.Name.Equals("ThisWorkbook", StringComparison.OrdinalIgnoreCase))
                    {
                        vbaModule = mod;
                        break;
                    }
                }

                if (vbaModule == null)
                {
                    int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Document, "ThisWorkbook");
                    vbaModule = workbook.VbaProject.Modules[moduleIndex];
                }

                // VBA code that runs when the workbook is opened.
                // It formats the range A1:B2: sets background color to yellow and makes the font bold.
                string vbaCode =
                    "Private Sub Workbook_Open()\r\n" +
                    "    Worksheets(\"Data\").Range(\"A1:B2\").Interior.Color = vbYellow\r\n" +
                    "    Worksheets(\"Data\").Range(\"A1:B2\").Font.Bold = True\r\n" +
                    "End Sub";

                vbaModule.Codes = vbaCode;

                // Save the workbook as a macro‑enabled file (.xlsm)
                string outputPath = "MacroEnabledWorkbook.xlsm";
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}