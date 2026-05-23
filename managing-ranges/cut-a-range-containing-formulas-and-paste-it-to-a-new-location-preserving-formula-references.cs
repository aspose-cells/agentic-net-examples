using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCutPasteDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate source range with values and formulas
                cells["A1"].PutValue(10);
                cells["A2"].Formula = "=A1*2";          // Formula referencing A1
                cells["B1"].PutValue(5);
                cells["B2"].Formula = "=B1+A2";         // Formula referencing B1 and A2

                // Define the range to cut (A1:B2)
                AsposeRange cutRange = cells.CreateRange("A1:B2");

                // Insert the cut range at a new location (starting at row 5, column 0 -> cell A6)
                // ShiftType.Down indicates that existing cells will be shifted down to make space
                cells.InsertCutCells(cutRange, 5, 0, ShiftType.Down);

                // Determine output file path
                string outputPath = "CutPasteFormulas.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}