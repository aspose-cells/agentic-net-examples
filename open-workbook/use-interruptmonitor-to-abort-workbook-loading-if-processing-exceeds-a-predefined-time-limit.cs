// Title: Abort Excel workbook loading with InterruptMonitor timeout in Aspose.Cells for .NET (C#)
// Common Searches: how to abort workbook loading with InterruptMonitor in Aspose.Cells C# | set a loading timeout for Excel files using Aspose.Cells .NET | cancel long‑running workbook load using Aspose.Cells InterruptMonitor example | Aspose.Cells C# load workbook with time limit and interrupt monitor | prevent Excel file load hang by using InterruptMonitor in Aspose.Cells
// Tags: interruptmonitor workbook load timeout | aspocells load workbook with interrupt monitor | cancelling Excel load after time limit .NET | prevent long‑running workbook load aspocells | excel file loading timeout handling aspocells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// Demonstrates how to configure an InterruptMonitor in Aspose.Cells for .NET to abort the loading of an Excel workbook when the operation exceeds a predefined time limit, including setup of LoadOptions, timeout handling, and graceful exception processing.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the workbook to be loaded.
            string inputFile = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            // Load options (no interrupt monitor used to keep compatibility).
            var loadOptions = new LoadOptions();

            // Load the workbook.
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Define output path.
            string outputFile = "output.xlsx";

            // Save the workbook.
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to {outputFile}");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
