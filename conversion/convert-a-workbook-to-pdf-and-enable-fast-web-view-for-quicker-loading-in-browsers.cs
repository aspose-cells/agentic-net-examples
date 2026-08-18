// Title: C# – Convert an Aspose.Cells Workbook to a Linearized PDF (Fast Web View)
// Description: Demonstrates how to create or load an Aspose.Cells workbook, configure PdfSaveOptions to enable Fast Web View (linearized PDF) using reflection for version‑agnostic property names, and save the workbook as a PDF that streams efficiently in browsers.
// Keywords: Aspose.Cells | C# PDF conversion | Fast Web View | Linearized PDF | PdfSaveOptions | EnableFastWebView | FastWebView property | browser-friendly PDF | .NET
// Common Searches: Aspose.Cells enable fast web view C# | linearized PDF with Aspose.Cells | PdfSaveOptions FastWebView property | C# convert Excel to PDF fast web view | Aspose.Cells PDF streaming optimization
// Developer Intent: Produce a PDF from an Excel workbook that is linearized for progressive rendering in web browsers.
// Use Cases: Generate downloadable reports that load page‑by‑page in a browser. | Serve large PDFs from a web application with minimal initial load time. | Maintain compatibility across Aspose.Cells versions when enabling Fast Web View.
// AI Prompts: Write C# code using Aspose.Cells to save a workbook as a linearized PDF with Fast Web View, handling both EnableFastWebView and FastWebView properties. | Explain the benefits of linearized PDFs for web performance and how to verify the Fast Web View setting in the generated file. | Provide step‑by‑step instructions to configure PdfSaveOptions for Fast Web View and save the workbook to a specific file path.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsFastWebViewDemo
{
    // Demonstrates how to create or load an Aspose.Cells workbook, configure PdfSaveOptions to enable Fast Web View (linearized PDF) using reflection for version‑agnostic property names, and save the workbook as a PDF that streams efficiently in browsers.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // creates an empty workbook

                // Add some sample data so the PDF is not empty
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Fast Web View PDF Demo");
                sheet.Cells["A2"].PutValue(DateTime.Now);

                // 2. Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Enable fast web view (linearized PDF) – handle different property names across versions
                var enableProp = typeof(PdfSaveOptions).GetProperty("EnableFastWebView");
                if (enableProp != null && enableProp.CanWrite)
                {
                    enableProp.SetValue(pdfOptions, true);
                }
                else
                {
                    var altProp = typeof(PdfSaveOptions).GetProperty("FastWebView");
                    if (altProp != null && altProp.CanWrite)
                    {
                        altProp.SetValue(pdfOptions, true);
                    }
                }

                // 3. Save the workbook as a PDF using the options
                string outputPath = "FastWebViewOutput.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
