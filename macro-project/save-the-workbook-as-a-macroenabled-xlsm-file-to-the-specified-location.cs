using System;
using Aspose.Cells;

namespace AsposeCellsMacroSaveDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default is a macro‑enabled workbook when saved as Xlsm)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Macro‑Enabled");
            sheet.Cells["B1"].PutValue(DateTime.Now);

            // Specify the output path
            string outputPath = "MacroEnabledWorkbook.xlsm";

            // Save the workbook as a macro‑enabled XLSM file
            // Uses the Save(string, SaveFormat) overload as defined in the Aspose.Cells API
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}