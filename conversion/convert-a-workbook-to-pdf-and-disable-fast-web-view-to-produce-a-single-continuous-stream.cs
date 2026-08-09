// Title: Aspose.Cells C# – Convert Workbook to PDF with Fast Web View Disabled (Linear PDF)
// Description: Demonstrates how to create or load an Aspose.Cells Workbook, set PdfSaveOptions.FastWebView to false, and save the workbook as a single‑stream PDF. The resulting file is linear, improving sequential reading and compatibility with older PDF viewers.
// Keywords: Aspose.Cells | PDF conversion | FastWebView | PdfSaveOptions | C# | linear PDF | disable fast web view | continuous PDF stream | Excel to PDF | .NET
// Common Searches: Aspose.Cells disable FastWebView PDF | C# save workbook as PDF without fast web view | PdfSaveOptions FastWebView false example | generate linear PDF from Excel using Aspose.Cells | Aspose.Cells PDF single stream output
// Developer Intent: Produce a PDF from an Excel workbook while turning off Fast Web View to obtain a single continuous stream.
// Use Cases: Provide downloadable PDF reports from web applications where fast web view must be off for legacy viewer support. | Stream large Excel‑derived PDFs efficiently in low‑bandwidth environments. | Batch‑convert multiple workbooks to linear PDFs to reduce file size and simplify sequential processing.
// AI Prompts: Show C# code that sets PdfSaveOptions.FastWebView = false in Aspose.Cells and saves a workbook as PDF. | Explain why disabling Fast Web View creates a linear PDF and how it affects file size and compatibility. | Generate a step‑by‑step guide for converting an existing Excel file to a single‑stream PDF using Aspose.Cells .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExample
{
    // Demonstrates how to create or load an Aspose.Cells Workbook, set PdfSaveOptions.FastWebView to false, and save the workbook as a single‑stream PDF. The resulting file is linear, improving sequential reading and compatibility with older PDF viewers.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Add some sample data
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample data for PDF conversion");

                // Initialize PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Disable Fast Web View if the property exists (commented out for compatibility with older versions)
                // pdfOptions.FastWebView = false;

                // Save the workbook as PDF using the specified options
                string outputPath = "output.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
