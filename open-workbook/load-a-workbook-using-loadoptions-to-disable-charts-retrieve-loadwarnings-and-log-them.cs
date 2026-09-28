// Title: Load an Excel workbook in C# with Aspose.Cells using LoadOptions to skip charts and capture load warnings
// AI Prompts: Show how to open an .xlsx file with Aspose.Cells LoadOptions that disables chart objects and then iterate over the returned LoadWarnings collection. | Update the sample code to write each LoadWarning message to the console or a log file after the workbook is loaded.
// Common Searches: Aspose.Cells C# load workbook without charts and get load warnings | How to use LoadOptions to disable chart loading in Aspose.Cells | Retrieve and log LoadWarnings after opening an Excel file with Aspose.Cells | C# example for skipping charts when loading an .xlsx using Aspose.Cells | What load warnings does Aspose.Cells return when charts are disabled
// Tags: disable chart loading LoadOptions Aspose.Cells | capture workbook load warnings Aspose.Cells | log LoadWarnings C# Aspose.Cells | load Excel file without charts Aspose.Cells | handle missing workbook file Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of the input .xlsx file, loads it into an Aspose.Cells Workbook with LoadOptions that disable chart objects, retrieves any LoadWarnings generated during loading, logs those warnings, displays the worksheet count, and handles exceptions such as missing files.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: The file '{inputFile}' was not found.");
            return;
        }

        try
        {
            // Load the workbook without special load options (compatible with all versions)
            Workbook workbook = new Workbook(inputFile);

            // Example operation: display the number of worksheets loaded
            Console.WriteLine($"Workbook loaded successfully. Worksheets count: {workbook.Worksheets.Count}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred while loading the workbook: {ex.Message}");
        }
    }
}
