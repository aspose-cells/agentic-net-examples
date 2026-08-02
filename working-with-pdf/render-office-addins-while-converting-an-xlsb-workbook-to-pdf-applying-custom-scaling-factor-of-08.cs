// Title: C# – Convert an Office Add‑In XLSB to PDF with 80% Scaling Using Aspose.Cells
// Description: Loads an Office Add‑In XLSB workbook, sets PageSetup.Zoom to 80 % on every sheet, saves a temporary file, and converts it to PDF with Aspose.Cells.Utility.ConversionUtility while preserving the custom scaling.
// Keywords: Aspose.Cells | C# | .NET | XLSB to PDF | Office Add‑In | page scaling | PageSetup.Zoom | ConversionUtility | PDF rendering | custom zoom
// Common Searches: Aspose.Cells convert XLSB to PDF with zoom | set worksheet zoom before PDF conversion C# | render Office Add‑In workbook as PDF using Aspose | apply 80% scaling to all sheets during PDF export | ConversionUtility PDF from modified XLSB
// Developer Intent: Generate a PDF from an Office Add‑In XLSB workbook while applying an 80 % page‑scaling factor to every worksheet.
// Use Cases: Produce printable PDFs from add‑in workbooks that fit standard paper sizes. | Batch‑process multiple XLSB add‑in files, applying a uniform zoom before archival PDF creation. | Create client‑ready PDFs where the original layout must be retained after scaling down content. | Automate report generation in CI pipelines with consistent page scaling.
// AI Prompts: Show C# code to set PageSetup.Zoom for all worksheets and convert the workbook to PDF with Aspose.Cells. | Give an example of robust error handling when using ConversionUtility to turn a temporary XLSB into PDF. | Explain how to calculate a dynamic zoom percentage based on worksheet dimensions before PDF export.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsAddInRendering
{
    // Loads an Office Add‑In XLSB workbook, sets PageSetup.Zoom to 80 % on every sheet, saves a temporary file, and converts it to PDF with Aspose.Cells.Utility.ConversionUtility while preserving the custom scaling.
    class Program
    {
        static void Main()
        {
            // Path to the source XLSB workbook (Office Add‑In file)
            string sourcePath = "AddInWorkbook.xlsb";

            // Path for the resulting PDF file
            string pdfPath = "AddInWorkbook.pdf";

            // Temporary file to store the workbook after applying page setup changes
            string tempPath = "TempModified.xlsb";

            try
            {
                // Verify that the source workbook exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {Path.GetFullPath(sourcePath)}");
                    return;
                }

                // Load the XLSB workbook
                Workbook workbook = new Workbook(sourcePath);

                // Apply a custom scaling factor of 80% to every worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.PageSetup.Zoom = 80; // 80% scaling
                }

                // Save the modified workbook to a temporary XLSB file
                workbook.Save(tempPath, new XlsbSaveOptions());

                // Convert the temporary XLSB file to PDF.
                // ConversionUtility respects the page setup (including Zoom) during conversion.
                ConversionUtility.Convert(tempPath, pdfPath);

                Console.WriteLine($"Conversion completed. PDF saved to: {Path.GetFullPath(pdfPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary file if it exists
                if (File.Exists(tempPath))
                {
                    try
                    {
                        File.Delete(tempPath);
                    }
                    catch
                    {
                        // Suppress any exceptions during cleanup
                    }
                }
            }
        }
    }
}
