// Title: Save names of Excel worksheets that contain only formatting (no data) to a text file with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find worksheets that have only formatting and no data | How to list Excel sheets without data rows using Aspose.Cells API | Export names of formatting‑only worksheets to a .txt file in .NET | Detect sheets that contain only styles, comments or shapes with Aspose.Cells | Save list of empty‑data Excel worksheets to a text file using C#
// Tags: worksheet formatting‑only detection Aspose.Cells | export worksheet names to txt C# | use Cells.MaxDataRow MaxDataColumn to check for data | list sheets without data rows Aspose.Cells | save formatting‑only sheet list .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorksheetFormattingOnlyExtractor
{
    // The example loads an Excel workbook, iterates through each worksheet, uses Cells.MaxDataRow and MaxDataColumn to identify sheets that contain no data (only formatting), collects their names, and writes the list to a specified text file.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: input Excel file path and output text file path.
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: WorksheetFormattingOnlyExtractor <input_excel_path> <output_text_path>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                List<string> formattingOnlySheets = new List<string>();

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Cells.MaxDataRow and MaxDataColumn return -1 when there is no cell containing data.
                    // If both are -1, the sheet contains only formatting (styles, comments, shapes, etc.).
                    if (sheet.Cells.MaxDataRow == -1 && sheet.Cells.MaxDataColumn == -1)
                    {
                        formattingOnlySheets.Add(sheet.Name);
                    }
                }

                // Write results to the output text file
                using (StreamWriter writer = new StreamWriter(outputPath, false))
                {
                    if (formattingOnlySheets.Count == 0)
                    {
                        writer.WriteLine("No worksheets with formatting‑only cells were found.");
                        Console.WriteLine("No worksheets with formatting‑only cells were found.");
                    }
                    else
                    {
                        writer.WriteLine("Worksheets containing only formatting cells:");
                        foreach (string name in formattingOnlySheets)
                        {
                            writer.WriteLine(name);
                        }

                        Console.WriteLine($"Found {formattingOnlySheets.Count} worksheet(s) with formatting‑only cells.");
                        Console.WriteLine($"Names have been saved to \"{outputPath}\".");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
