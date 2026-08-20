// Title: C# – Convert Excel to PDF with a 20‑second timeout using Aspose.Cells InterruptMonitor
// Description: Loads an XLSX file, attaches a SystemTimeInterruptMonitor via LoadOptions, starts a 20 000 ms timer, and saves the workbook as PDF. If the export exceeds the limit, a CellsException with the Interrupted code is thrown and handled.
// Keywords: Aspose.Cells | C# Excel to PDF | InterruptMonitor | SystemTimeInterruptMonitor | timeout PDF conversion | abort long export | Workbook.Save timeout | SaveFormat.Pdf | performance safeguard | large Excel files
// Common Searches: Aspose.Cells set timeout for PDF export C# | How to abort Excel to PDF conversion after 20 seconds | SystemTimeInterruptMonitor example for workbook.Save | C# limit Aspose.Cells PDF generation time | Catch CellsException.Interrupted during save
// Developer Intent: Create a PDF conversion that automatically stops when processing exceeds twenty seconds.
// Use Cases: Prevent web‑service requests from hanging while converting large spreadsheets to PDF. | Add a safety guard in batch jobs that generate PDFs to avoid server timeouts. | Provide immediate feedback to users when a conversion is terminated due to time constraints.
// AI Prompts: Write C# code that uses Aspose.Cells and SystemTimeInterruptMonitor to export an Excel workbook to PDF with a 15‑second limit. | Explain how to detect and handle the CellsException.Interrupted error during a workbook.Save operation. | Show how to reuse a single InterruptMonitor for both loading and saving in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads an XLSX file, attaches a SystemTimeInterruptMonitor via LoadOptions, starts a 20 000 ms timer, and saves the workbook as PDF. If the export exceeds the limit, a CellsException with the Interrupted code is thrown and handled.
class Program
{
    static void Main()
    {
        // Input workbook file (replace with actual path)
        string inputPath = "input.xlsx";
        // Output PDF file
        string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Create a SystemTimeInterruptMonitor that throws an exception when interrupted
        SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);

        // Attach the monitor to load options so it is active during loading
        LoadOptions loadOptions = new LoadOptions
        {
            InterruptMonitor = monitor
        };

        Workbook workbook;
        try
        {
            // Load the workbook with the interrupt monitor attached
            workbook = new Workbook(inputPath, loadOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Assign the monitor to the workbook to monitor the save operation
        workbook.InterruptMonitor = monitor;

        // Start the monitor with a 20‑second (20000 ms) limit before saving
        monitor.StartMonitor(20000);

        try
        {
            // Save the workbook as PDF; will be aborted if it exceeds 20 seconds
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine("Workbook successfully saved to PDF.");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            Console.WriteLine("Save operation was interrupted after exceeding the time limit.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error during save: {ex.Message}");
        }
        finally
        {
            // No explicit StopMonitor method; monitor will be disposed automatically when out of scope
        }
    }
}
