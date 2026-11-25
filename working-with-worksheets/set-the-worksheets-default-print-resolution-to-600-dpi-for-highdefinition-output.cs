// Title: Set worksheet default print resolution to 600 DPI for PDF export using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set default image resolution to 600 DPI for PDF output | How to increase PDF export resolution to 600 DPI with Aspose.Cells workbook | C# change worksheet print quality to 600 DPI using Aspose.Cells settings | Set high‑definition print resolution in Aspose.Cells before saving as PDF | DefaultImageResolution property 600 DPI Aspose.Cells example
// Tags: configure workbook DPI using Aspose.Cells | high‑definition PDF generation C# Aspose.Cells | worksheet page setup for print quality | Aspose.Cells image resolution property example | export workbook to PDF with 600 DPI

using System;
using Aspose.Cells;

namespace SetPrintResolutionDemo
{
    // The sample creates a workbook, adds data, configures page orientation and paper size, sets the workbook's DefaultImageResolution to 600 DPI, and saves the worksheet as a high‑definition PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // OPTIONAL: Set Aspose.Cells license if you have one.
                // var license = new Aspose.Cells.License();
                // license.SetLicense("Aspose.Cells.lic");

                // Create a new workbook and access the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Name = "HighDefSheet";

                // Add sample data
                worksheet.Cells["A1"].PutValue("High‑definition print resolution demo");
                worksheet.Cells["A2"].PutValue(DateTime.Now);
                worksheet.Cells["A3"].PutValue("Resolution set to 600 DPI");

                // Adjust page‑setup settings
                worksheet.PageSetup.Orientation = PageOrientationType.Portrait;
                worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
                worksheet.PageSetup.FitToPagesWide = 1;
                worksheet.PageSetup.FitToPagesTall = 0; // auto height

                // Set default image (and PDF) resolution to 600 DPI
                workbook.Settings.DefaultImageResolution = 600;

                // Save the workbook to PDF
                string outputPath = "HighResolutionOutput.pdf";
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\" with print resolution set to 600 DPI.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
