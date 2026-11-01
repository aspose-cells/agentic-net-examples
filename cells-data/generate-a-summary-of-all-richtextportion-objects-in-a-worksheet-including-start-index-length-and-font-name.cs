// Title: Iterate through used cells in the first worksheet and list each string cell’s text with its font name using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loops over the used range of the first worksheet and prints the cell address, its string value, and the font name applied to the whole cell. | Enhance the example to retrieve every RichTextPortion from string cells and output each portion’s start index, length, and font name. | Change the program so that the generated summary lines are written to a new worksheet instead of being displayed on the console.
// Common Searches: asp.net core extract font name of each string cell using Aspose.Cells | c# list all richtext portions in an Excel worksheet with start index and length | how to get cell text and style font from used range Aspose.Cells C# | iterate over used cells and output rich text portion details in Aspose.Cells | save summary of cell contents and fonts to new sheet Aspose.Cells .NET
// Tags: Aspose.Cells enumerate used cells C# | retrieve cell string value and font name | extract RichTextPortion start index length Aspose.Cells | write summary to new worksheet Aspose.Cells | C# Excel cell style extraction Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRichTextSummary
{
    // The sample loads an Excel workbook, walks through every used cell in the first worksheet, captures the text and the font name for cells containing strings, prints a formatted summary line for each, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "Input.xlsx";
                const string outputPath = "Output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load workbook
                var workbook = new Workbook(inputPath);
                var worksheet = workbook.Worksheets[0];
                var cells = worksheet.Cells;

                var summaries = new List<string>();

                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                // Iterate through used cells
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        var cell = cells[row, col];

                        // Process only string cells
                        if (cell.Type == CellValueType.IsString)
                        {
                            try
                            {
                                // Retrieve cell text
                                string text = cell.StringValue;

                                // Retrieve the font name applied to the whole cell
                                string fontName = cell.GetStyle().Font?.Name ?? "Default";

                                string summary = $"Cell[{row},{col}] Text=\"{text}\", FontName={fontName}";
                                summaries.Add(summary);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error processing cell [{row},{col}]: {ex.Message}");
                            }
                        }
                    }
                }

                // Output the summary
                foreach (var line in summaries)
                {
                    Console.WriteLine(line);
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
