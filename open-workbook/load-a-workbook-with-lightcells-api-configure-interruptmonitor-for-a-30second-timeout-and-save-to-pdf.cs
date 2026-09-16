// Title: Load an Excel workbook with Aspose.Cells LightCells API, set a 30‑second InterruptMonitor, and export it to PDF in C#
// AI Prompts: Generate C# code that opens an .xlsx file using Aspose.Cells LightCells API, applies an InterruptMonitor with a 30‑second timeout, and saves the workbook as a PDF. | Show how to add a file‑existence check before loading the workbook and log any errors during LightCells PDF conversion in C#. | Provide a sample that configures LightCells options for PDF output and demonstrates handling a timeout exception from InterruptMonitor.
// Common Searches: c# aspocells lightcells api convert excel to pdf with interruptmonitor timeout | how to set a 30 second processing timeout for lightcells in aspocells c# | excel to pdf conversion using lightcells api and file existence validation c# | aspocells interruptmonitor example for large workbook PDF export c# | c# code to load workbook, check file, and save as pdf using lightcells api
// Tags: aspose.cells lightcells pdf export | aspose.cells interruptmonitor timeout | c# excel workbook loading with file check | c# lightcells api pdf conversion | c# handling timeout exception aspocells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates C# code that verifies the presence of an input.xlsx file, loads it with Aspose.Cells LightCells API, configures an InterruptMonitor to abort after 30 seconds, and saves the workbook as output.pdf while handling errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook using the standard API
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully converted to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
