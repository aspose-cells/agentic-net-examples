using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsAutoFitDemo
{
    public class AutoFitColumnsExample
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                worksheet.Cells["A1"].PutValue("ID");
                worksheet.Cells["B1"].PutValue("Product Description");
                worksheet.Cells["C1"].PutValue("Price");

                worksheet.Cells["A2"].PutValue(101);
                worksheet.Cells["B2"].PutValue("Compact widget");
                worksheet.Cells["C2"].PutValue(12.5);

                worksheet.Cells["A3"].PutValue(102);
                worksheet.Cells["B3"].PutValue("Advanced multi-purpose widget with extended features and a very long description");
                worksheet.Cells["C3"].PutValue(199.99);

                // Auto‑fit all columns so each column width matches its longest cell content
                worksheet.AutoFitColumns();

                // Save the workbook
                string outputPath = "AutoFitColumnsResult.xlsx";
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
            AutoFitColumnsExample.Run();
        }
    }
}