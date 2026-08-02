// Title: Toggle slicer visibility in PDF export using Aspose.Cells for .NET
// Description: Shows how a boolean setting can control the printable state of slicers, letting you include or omit them when converting an Excel workbook to PDF with Aspose.Cells.
// Keywords: Aspose.Cells | C# | PDF export | slicer visibility | Shape.IsPrintable | conditional export | runtime flag | Excel to PDF | hide slicer | Aspose.Cells slicer
// Common Searches: Aspose.Cells hide slicer in PDF | set slicer printable property C# | conditional slicer export Aspose | toggle slicer visibility PDF Aspose.Cells | read config flag to control slicer PDF
// Developer Intent: Needs to decide at runtime whether slicers are rendered in the generated PDF file.
// Use Cases: Create a PDF report that retains slicer controls for interactive review. | Generate a clean PDF without UI elements by disabling slicer printing. | Read a boolean from appsettings.json and apply it to all slicers before conversion. | Assign different printable states to individual slicers in a multi‑sheet workbook.
// AI Prompts: Write C# code that reads a boolean from appsettings.json and sets slicer.Shape.IsPrintable before saving the workbook to PDF with PdfSaveOptions. | Explain how Shape.IsPrintable influences slicer rendering in PDF output and how to toggle it based on user input. | Provide a sample that iterates over every slicer in a workbook and applies visibility flags from a configuration dictionary prior to PDF export.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;
using Aspose.Cells.Rendering;

namespace AsposeCellsSlicerPdfExport
{
    // Shows how a boolean setting can control the printable state of slicers, letting you include or omit them when converting an Excel workbook to PDF with Aspose.Cells.
    class Program
    {
        static void Main()
        {
            try
            {
                // Configuration flag: set to true to keep slicers in the exported PDF, false to hide them.
                bool keepSlicers = true; // Could be read from a config file or command‑line argument.

                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate some sample data for the slicer to work with.
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("A");
                worksheet.Cells["A5"].PutValue("B");

                // Convert the range into a table (ListObject) which slicers can be attached to.
                int tableIndex = worksheet.ListObjects.Add(0, 0, 0, 0, true);
                ListObject table = worksheet.ListObjects[tableIndex];
                // Use DisplayName instead of Name (Name property not available in this version).
                table.DisplayName = "CategoryTable";

                // Add a slicer linked to the table.
                int slicerIndex = worksheet.Slicers.Add(table, 0, "D1");
                Slicer slicer = worksheet.Slicers[slicerIndex];

                // Use the configuration flag to decide whether the slicer should be printable.
                // The Shape.IsPrintable property determines if the slicer appears in the PDF output.
                slicer.Shape.IsPrintable = keepSlicers;

                // Prepare PDF save options (default options are sufficient for this example).
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook to PDF using the configured options.
                string outputPath = "WorkbookWithSlicer.pdf";
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF saved to '{outputPath}'. Slicers kept: {keepSlicers}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
