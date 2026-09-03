// Title: Save names of formatting‑only worksheets to a text file with Aspose.Cells for .NET
// AI Prompts: Write a C# console program that uses Aspose.Cells to load a workbook, finds worksheets where MaxDataRow equals –1, and writes each worksheet name on a separate line to a given text file. | Create a .NET utility that accepts an Excel file path and an output .txt path, scans all sheets for the absence of data rows using Aspose.Cells, and records the names of those formatting‑only sheets.
// Common Searches: Aspose.Cells find worksheets that contain only formatting and export their names to a txt file | C# list Excel sheets with no data rows using Aspose.Cells | How to detect formatting‑only worksheets in a workbook with Aspose.Cells .NET | Save names of empty or formatting‑only Excel sheets to a text file in C# | Console application to write worksheet names without values to a file using Aspose.Cells
// Tags: detect formatting‑only worksheets Aspose.Cells | export worksheet names to txt C# | maxdatarow check Aspose.Cells | list sheets without data rows .NET | write worksheet list to text file Aspose.Cells | identify formatting‑only Excel sheets C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads a workbook, iterates through each worksheet, uses the MaxDataRow property to identify sheets that have no data (only formatting), collects those sheet names, and writes them line‑by‑line to a specified text file.
class Program
{
    static void Main(string[] args)
    {
        // Validate arguments: args[0] = input workbook path, args[1] = output text file path
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: Program.exe <inputWorkbookPath> <outputTextFilePath>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(inputPath);

        // List to hold names of worksheets that contain only formatting (no data)
        List<string> formattingOnlySheets = new List<string>();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // MaxDataRow returns -1 if there are no cells with data (values or formulas)
            // This effectively identifies sheets that have only formatting or are completely empty
            if (sheet.Cells.MaxDataRow == -1)
            {
                formattingOnlySheets.Add(sheet.Name);
            }
        }

        // Write the collected worksheet names to the output text file, one name per line
        File.WriteAllLines(outputPath, formattingOnlySheets);

        Console.WriteLine($"Found {formattingOnlySheets.Count} worksheet(s) with only formatting.");
        Console.WriteLine($"Worksheet names saved to: {outputPath}");
    }
}
