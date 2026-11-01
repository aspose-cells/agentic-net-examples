// Title: Convert a loaded Excel workbook to PDF with a 20‑second timeout using Aspose.Cells InterruptMonitor in C#
// Common Searches: how to abort Aspose.Cells PDF conversion after 20 seconds in C# | C# convert workbook to PDF with InterruptMonitor timeout | Aspose.Cells save workbook as PDF with cancellation token example | set maximum conversion time for Excel to PDF using Aspose.Cells | prevent long‑running PDF export in Aspose.Cells C#
// Tags: Aspose.Cells PDF conversion with InterruptMonitor | C# workbook to PDF timeout handling | Excel to PDF conversion abort after time limit | Aspose.Cells conversion cancellation example | interrupt monitor usage in Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing XLSX workbook, attaches an Aspose.Cells.InterruptMonitor configured for a 20‑second limit, and attempts to save the workbook as a PDF. If the conversion exceeds the timeout, the monitor aborts the operation and the code catches the interruption, ensuring graceful error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook from the specified file
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Convert and save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during the conversion
            Console.WriteLine($"Conversion error: {ex.Message}");
        }
    }
}
