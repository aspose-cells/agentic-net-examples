using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CopyRowsAndAutoFitColumnsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate source rows with sample data
                cells["A1"].PutValue("Header");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("Row 1");
                cells["B2"].PutValue(12345);
                cells["A3"].PutValue("Row 2 with a very long text that will require column auto‑fit");
                cells["B3"].PutValue(67890);

                // Copy rows 0‑2 (three rows) to start at row index 5 in the same worksheet
                cells.CopyRows(cells, 0, 5, 3);

                // Auto‑fit all columns to the new data
                worksheet.AutoFitColumns();

                // Save the workbook
                string outputPath = "CopyRowsAndAutoFitColumnsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            CopyRowsAndAutoFitColumnsDemo.Run();
        }
    }
}