// Title: Log original and updated worksheet page‑setup properties (paper size and orientation) to a text file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, iterates through every worksheet, sets each sheet’s PaperSize to PaperA4, and writes the previous and new PaperSize and Orientation values to a log file. | Extend the Aspose.Cells example to also capture the Top, Bottom, Left, and Right margin values and output the complete audit information in CSV format. | Add robust error handling that catches exceptions during PageSetup access or workbook saving and records any failures in the same log file.
// Common Searches: Aspose.Cells C# log worksheet page setup changes to a file | how to record original paper size and orientation when modifying Excel sheets with Aspose.Cells | C# write Excel page setup audit log for each worksheet using Aspose.Cells
// Tags: Aspose.Cells log page setup changes | C# record worksheet paper size Aspose | Aspose.Cells audit worksheet page orientation | write Excel page setup audit log C#

using System;
using System.IO;
using Aspose.Cells;

// The example opens an existing workbook, iterates over all worksheets, changes each sheet’s paper size to A4, logs the previous and new paper size and orientation to PageSetupChanges.log, and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare the log file (overwrites if it already exists)
            string logPath = "PageSetupChanges.log";
            using (StreamWriter logWriter = new StreamWriter(logPath, false))
            {
                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Capture the original page‑setup settings
                    PageSetup oldSetup = sheet.PageSetup;
                    PaperSizeType oldPaperSize = oldSetup.PaperSize;
                    string oldOrientation = oldSetup.Orientation.ToString();

                    // ----- BEGIN PAGE‑SETUP MODIFICATION -----
                    // Example change: set paper size to A4
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
                    // ----- END PAGE‑SETUP MODIFICATION -----

                    // Capture the new page‑setup settings after modification
                    PageSetup newSetup = sheet.PageSetup;
                    PaperSizeType newPaperSize = newSetup.PaperSize;
                    string newOrientation = newSetup.Orientation.ToString();

                    // Log detailed information about the change
                    logWriter.WriteLine($"Worksheet: {sheet.Name}");
                    logWriter.WriteLine($"  PaperSize   : {oldPaperSize} -> {newPaperSize}");
                    logWriter.WriteLine($"  Orientation : {oldOrientation} -> {newOrientation}");
                    logWriter.WriteLine(); // Blank line for readability
                }
            }

            // Save the modified workbook to a new file
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
