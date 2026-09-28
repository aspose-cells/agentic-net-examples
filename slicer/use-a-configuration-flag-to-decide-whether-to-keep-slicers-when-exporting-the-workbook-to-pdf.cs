// Title: Export an Aspose.Cells workbook to PDF while optionally preserving slicers using a configuration flag in C#
// AI Prompts: Write C# code that creates a worksheet, adds a table and a linked slicer, reads a boolean flag from app settings or command line, sets the slicer's Shape.IsPrintable property based on that flag, and saves the workbook as a PDF with Aspose.Cells. | Show how to toggle slicer visibility during PDF export in Aspose.Cells by controlling the printable attribute of the slicer shape according to a configuration value. | Provide a complete example that demonstrates reading a configuration setting, applying it to a slicer, and exporting the workbook to PDF using PdfSaveOptions in .NET.
// Common Searches: Aspose.Cells C# export workbook to PDF keep slicer based on app setting | how to hide slicer in PDF output with Aspose.Cells | set slicer printable flag before saving as PDF in .NET | conditional PDF export of slicers using a configuration flag in Aspose.Cells
// Tags: Aspose.Cells PDF export slicer control | slicer printable property Aspose.Cells | conditional slicer visibility PDF Aspose.Cells | C# read config flag Aspose.Cells | export workbook to PDF using PdfSaveOptions

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;
using Aspose.Cells.Tables;   // Required for ListObject

namespace AsposeCellsSlicerPdfExport
{
    // The example creates a workbook, adds a table and a linked slicer, reads a boolean configuration flag, sets the slicer's printable property accordingly, and saves the workbook as a PDF using Aspose.Cells PdfSaveOptions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Configuration flag: true to keep slicers in the PDF, false to hide them
                bool keepSlicers = true; // This could be read from config, command‑line, etc.

                // Create a new workbook and add some data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                // Create a table from the data (required for slicer)
                int tableIndex = sheet.ListObjects.Add(0, 0, 3, 0, true);
                ListObject table = sheet.ListObjects[tableIndex];
                // Set a display name for the table (Name property may not be available in some versions)
                table.DisplayName = "CategoryTable";

                // Add a slicer linked to the table
                int slicerIndex = sheet.Slicers.Add(table, 0, "D1");
                Slicer slicer = sheet.Slicers[slicerIndex];

                // Set slicer printable flag based on the configuration
                slicer.Shape.IsPrintable = keepSlicers;

                // Configure PDF save options (default options are sufficient for this demo)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Determine output path and ensure directory exists
                string outputPath = "WorkbookWithSlicer.pdf";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to PDF
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook saved to PDF. Slicers kept: {keepSlicers}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
