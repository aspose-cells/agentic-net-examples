// Title: Lock All Shapes Across Worksheets and Export to PDF with Aspose.Cells for .NET
// Description: C# example that loads a workbook, iterates through every worksheet, sets each shape's IsLocked property to true, applies worksheet protection, and saves the result as a PDF using Aspose.Cells.
// Keywords: Aspose.Cells lock shapes | protect worksheet C# | export workbook to PDF | batch shape locking .NET | Aspose.Cells PDF conversion | shape security Aspose.Cells
// Common Searches: how to lock all shapes in an Excel file using Aspose.Cells | C# batch lock shapes before PDF export | protect worksheets and shapes with Aspose.Cells .NET | export protected workbook to PDF Aspose.Cells
// Developer Intent: Secure every shape in all worksheets and generate a PDF of the protected workbook.
// Use Cases: Secure financial dashboards by locking chart objects before distributing PDFs. | Preserve template integrity in multi‑sheet reports when archiving as PDF. | Enforce shape protection in regulatory filings generated from Excel workbooks.
// AI Prompts: Write C# code that locks every shape on each worksheet, protects the sheets, and saves the workbook as a PDF using Aspose.Cells. | Show how to customize PdfSaveOptions while keeping all shapes locked in an Aspose.Cells .NET project.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsBatchLockAndExport
{
    // C# example that loads a workbook, iterates through every worksheet, sets each shape's IsLocked property to true, applies worksheet protection, and saves the result as a PDF using Aspose.Cells.
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Lock every shape on the current worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    shape.IsLocked = true;
                }

                // Protect the worksheet so that the locked state takes effect
                sheet.Protect(ProtectionType.All);
            }

            // Prepare PDF save options (default options are sufficient for this task)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Export the workbook to PDF
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}
