using System;
using Aspose.Cells;

namespace AsposeCellsMacroSaveDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path where the macro‑enabled workbook will be saved
            string outputPath = "MacroEnabledWorkbook.xlsm";

            // Create a new workbook (you can also load an existing one)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Save the workbook as a macro‑enabled XLSM file
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
    }
}