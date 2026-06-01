using System;
using System.IO;
using Aspose.Cells;

namespace MarketingBrochure
{
    public class DisableRowHeadings
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Price");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(2.5);
                worksheet.Cells["A3"].PutValue("Orange");
                worksheet.Cells["B3"].PutValue(1.8);

                // Disable printing of row and column headings for a cleaner layout
                worksheet.PageSetup.PrintHeadings = false;

                // Optionally set the print area
                worksheet.PageSetup.PrintArea = "A1:B3";

                // Define output file path
                string outputPath = "Brochure.pdf";

                // Save the workbook as PDF for the marketing brochure
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DisableRowHeadings.Run();
        }
    }
}