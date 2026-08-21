// Title: Capture Font Substitution Warnings When Converting Excel to PDF with Aspose.Cells for .NET
// Description: Demonstrates how to implement a custom IWarningCallback to record only FontSubstitution warnings during SheetRender rendering, log each warning to the console, count them, and then save the workbook as a PDF using PdfSaveOptions.
// Keywords: Aspose.Cells | C# | .NET | font substitution warning | IWarningCallback | ImageOrPrintOptions WarningCallback | SheetRender | PdfSaveOptions | log rendering warnings | capture font warnings
// Common Searches: Aspose.Cells capture font substitution warnings | C# IWarningCallback example for PDF conversion | how to log missing font warnings in Aspose.Cells | retrieve font substitution warnings during Excel to PDF | custom warning callback Aspose.Cells .NET
// Developer Intent: The developer needs to detect and log any font substitution events that occur while rendering an Excel worksheet and saving it as a PDF.
// Use Cases: Identify missing fonts before final PDF generation and prompt the user to install or replace them. | Store font substitution details in an audit log for compliance reporting. | Display the number of font warnings in a UI dashboard after conversion completes.
// AI Prompts: Show a C# snippet that uses IWarningCallback to capture font substitution warnings during PDF export with Aspose.Cells. | Explain how to filter WarningInfo objects for FontSubstitution type and write the messages to a log file. | Provide code to replace missing fonts automatically after detecting substitution warnings in Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsFontSubstitutionDemo
{
    // Custom warning callback to capture font substitution warnings
    // Demonstrates how to implement a custom IWarningCallback to record only FontSubstitution warnings during SheetRender rendering, log each warning to the console, count them, and then save the workbook as a PDF using PdfSaveOptions.
    public class RenderingWarningCallback : IWarningCallback
    {
        // Store captured warnings
        public List<WarningInfo> CapturedWarnings { get; } = new List<WarningInfo>();

        // Called by Aspose.Cells during rendering
        public void Warning(WarningInfo warningInfo)
        {
            // Capture only font substitution warnings
            if (warningInfo.Type == ExceptionType.FontSubstitution)
            {
                CapturedWarnings.Add(warningInfo);
                Console.WriteLine($"Font substitution warning: {warningInfo.Description}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add text that uses a font unlikely to exist on the system
                Cell cell = sheet.Cells["A1"];
                cell.PutValue("Text with a non‑existent font");

                // Apply the non‑existent font to the cell
                Style style = workbook.CreateStyle();
                style.Font.Name = "NonExistentFont";
                cell.SetStyle(style);

                // Prepare rendering options and assign the custom warning callback
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    // Default image format is PNG; explicit setting omitted to avoid API mismatch
                    WarningCallback = new RenderingWarningCallback()
                };

                // Render the worksheet to an image (output not needed)
                SheetRender renderer = new SheetRender(sheet, renderOptions);
                using (MemoryStream dummyStream = new MemoryStream())
                {
                    try
                    {
                        renderer.ToImage(0, dummyStream);
                    }
                    catch (Exception renderEx)
                    {
                        Console.WriteLine($"Rendering error: {renderEx.Message}");
                    }
                }

                // Retrieve the callback to inspect captured warnings
                var warningCallback = (RenderingWarningCallback)renderOptions.WarningCallback;
                Console.WriteLine($"Total font substitution warnings captured: {warningCallback.CapturedWarnings.Count}");

                // Save the workbook to PDF (warnings already captured)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Use a common font to avoid further substitution
                    DefaultFont = "Arial"
                };

                string outputPath = "Output.pdf";
                try
                {
                    workbook.Save(outputPath, pdfOptions);
                    Console.WriteLine($"Workbook saved to PDF: {Path.GetFullPath(outputPath)}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving PDF: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
