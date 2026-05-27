using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetZoomAndExportPdf
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set the page setup zoom to 80%
                worksheet.PageSetup.Zoom = 80;
                // Ensure the scaling mode uses percent scale
                worksheet.PageSetup.IsPercentScale = true;

                // Add sample data to visualize in the PDF
                worksheet.Cells["A1"].PutValue("Sample Data");
                worksheet.Cells["B1"].PutValue(123);
                worksheet.Cells["A2"].PutValue("More Data");
                worksheet.Cells["B2"].PutValue(456);

                // Define output PDF path
                string outputPath = "Worksheet_Zoom_80.pdf";

                // Save the workbook as PDF
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
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
            SetZoomAndExportPdf.Run();
        }
    }
}