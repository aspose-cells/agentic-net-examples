using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SalesDataWorkbook
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (default Xlsx format)
                Workbook workbook = new Workbook();

                // Access the default first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SalesData";

                // Add header row
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Product");
                sheet.Cells["C1"].PutValue("Units Sold");
                sheet.Cells["D1"].PutValue("Revenue");

                // Populate sample sales data
                sheet.Cells["A2"].PutValue("January");
                sheet.Cells["B2"].PutValue("Widget");
                sheet.Cells["C2"].PutValue(120);
                sheet.Cells["D2"].PutValue(2400);

                sheet.Cells["A3"].PutValue("February");
                sheet.Cells["B3"].PutValue("Widget");
                sheet.Cells["C3"].PutValue(150);
                sheet.Cells["D3"].PutValue(3000);

                sheet.Cells["A4"].PutValue("January");
                sheet.Cells["B4"].PutValue("Gadget");
                sheet.Cells["C4"].PutValue(80);
                sheet.Cells["D4"].PutValue(2000);

                sheet.Cells["A5"].PutValue("February");
                sheet.Cells["B5"].PutValue("Gadget");
                sheet.Cells["C5"].PutValue(95);
                sheet.Cells["D5"].PutValue(2375);

                // Define output file path
                string outputPath = "SalesData.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SalesDataWorkbook.Run();
        }
    }
}