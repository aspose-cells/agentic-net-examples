using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // For PivotFieldType
using Aspose.Cells.Slicers; // For slicer support (optional)

namespace SlicerPrintableDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                var cells = sheet.Cells;

                // Populate sample data for the pivot table
                cells["A1"].Value = "Category";
                cells["A2"].Value = "A";
                cells["A3"].Value = "B";
                cells["A4"].Value = "C";

                cells["B1"].Value = "Value";
                cells["B2"].Value = 10;
                cells["B3"].Value = 20;
                cells["B4"].Value = 30;

                // Add a pivot table based on the data range
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "MyPivot");
                var pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");

                // Add a slicer linked to the pivot table
                int slicerIdx = sheet.Slicers.Add(pivot, "E1", "Category");
                var slicer = sheet.Slicers[slicerIdx];

                // Make the slicer printable so it appears on printed pages
                slicer.Shape.IsPrintable = true;

                // Define output file path
                string outputPath = "SlicerPrintableDemo.xlsx";

                // Ensure the output directory exists (prevents FileNotFoundException on save)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}