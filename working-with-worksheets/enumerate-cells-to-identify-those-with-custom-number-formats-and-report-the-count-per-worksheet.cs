// Title: Count cells with custom number formats on each worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to count cells that have custom number formats per worksheet | How to enumerate all cells in a workbook and detect custom number format styles with Aspose.Cells | Report the number of cells using user‑defined number formats for each sheet in an Excel file using Aspose.Cells .NET | Aspose.Cells C# code to list worksheets and tally cells with non‑built‑in number formats | Command‑line tool to find custom formatted cells in an Excel workbook with Aspose.Cells
// Tags: custom number format counting Aspose.Cells | enumerate worksheet cells C# Aspose.Cells | detect user‑defined number formats Excel .NET | cell style Custom property Aspose.Cells | command line workbook processing Aspose.Cells | per‑sheet custom format statistics .NET

using System;
using Aspose.Cells;

namespace CustomNumberFormatCounter
{
    // A C# console application that loads an Excel workbook, iterates through each worksheet, enumerates every cell, checks the Style.Custom property to identify user‑defined number formats, counts such cells, and prints the count per worksheet. Includes command‑line argument validation, file‑existence checks, and exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: CustomNumberFormatCounter <input-workbook-path>");
                return;
            }

            string inputPath = args[0];

            if (!System.IO.File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File not found - {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Workbook loaded: {inputPath}");
                Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");
                Console.WriteLine();

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int customFormatCount = 0;

                    // Enumerate all cells in the worksheet
                    foreach (Cell cell in sheet.Cells)
                    {
                        // Retrieve the style of the cell
                        Style style = cell.GetStyle();

                        // The Custom property returns an empty string if the format is built‑in.
                        // A non‑empty string indicates a custom number format.
                        if (!string.IsNullOrEmpty(style.Custom))
                        {
                            customFormatCount++;
                        }
                    }

                    Console.WriteLine($"Worksheet \"{sheet.Name}\": {customFormatCount} cell(s) with custom number formats.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
