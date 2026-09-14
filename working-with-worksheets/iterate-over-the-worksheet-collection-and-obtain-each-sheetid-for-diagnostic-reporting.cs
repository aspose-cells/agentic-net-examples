// Title: How to loop through all worksheets in an Excel workbook and get each worksheet's Index (SheetId) using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, iterates over workbook.Worksheets, and prints each worksheet's Name together with its zero‑based Index. | Create a C# console program that uses Aspose.Cells to enumerate every worksheet in a workbook and output the sheet's diagnostic identifier for logging.
// Common Searches: C# Aspose.Cells get worksheet index for each sheet in a workbook | list worksheet IDs in an Excel file using Aspose.Cells .NET | how to display sheet order numbers with Aspose.Cells in C# | diagnostic printing of worksheet indexes Aspose.Cells example | retrieve zero based sheet index Aspose.Cells C# loop
// Tags: Aspose.Cells enumerate worksheets C# | retrieve worksheet Index property Aspose.Cells | Excel workbook sheet IDs Aspose.Cells | diagnostic worksheet enumeration .NET | loop over workbook worksheets Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// // Loads an Excel workbook (creating a default one if missing), iterates through each worksheet using Aspose.Cells, writes the worksheet name and its zero‑based Index to the console for diagnostics, and saves the workbook to an output file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                // Create a new workbook with a default worksheet
                var newWorkbook = new Workbook();
                newWorkbook.Worksheets[0].Name = "Sheet1";
                newWorkbook.Save(inputPath);
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet and retrieve its index (zero‑based)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int sheetIndex = sheet.Index; // Diagnostic identifier
                Console.WriteLine($"Worksheet \"{sheet.Name}\" has Index: {sheetIndex}");
            }

            // Save the workbook (even if unchanged) to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
