// Title: C# – Replace “Draft” with “Final” in a Named Range using Aspose.Cells for .NET
// Description: Loads an Excel workbook, retrieves a defined named range (e.g., PublishRange), parses its RefersTo formula to identify the worksheet and address, iterates every cell in the range, swaps any occurrence of the word “Draft” with “Final”, and saves the modified file. Includes checks for missing files, undefined names, and malformed formulas.
// Keywords: Aspose.Cells | C# replace text | named range | Excel workbook | .NET | PublishRange | bulk cell update | replace Draft with Final | cell iteration | RefersTo parsing | Excel automation
// Common Searches: replace text in named range Aspose.Cells | C# change all Draft cells to Final in Excel | how to get RefersTo address from named range .NET | iterate over cells in a specific range using Aspose.Cells | update placeholder values in Excel with Aspose.Cells | error handling for missing named range Aspose.Cells
// Developer Intent: Swap every occurrence of the word “Draft” for “Final” inside a specific named range before publishing the workbook.
// Use Cases: Finalize status labels in a publishing range prior to report distribution | Automate cleanup of draft markers across a predefined area of a spreadsheet | Prepare a template workbook for external stakeholders by converting draft tags to final tags | Integrate into a CI pipeline that validates and finalizes Excel assets | Support multi‑sheet workbooks where only a named range needs text replacement
// AI Prompts: Generate C# code with Aspose.Cells that finds a named range, parses its RefersTo property, iterates all cells, and replaces the substring 'Draft' with 'Final'. | Explain how to safely handle missing named ranges and malformed RefersTo formulas when performing bulk text replacement in Aspose.Cells. | Provide a step‑by‑step guide to bulk‑replace a word in a defined Excel range using Aspose.Cells for .NET, including error handling and performance tips.

using System;
using System.IO;
using Aspose.Cells;

namespace ReplaceDraftInNamedRange
{
    // Loads an Excel workbook, retrieves a defined named range (e.g., PublishRange), parses its RefersTo formula to identify the worksheet and address, iterates every cell in the range, swaps any occurrence of the word “Draft” with “Final”, and saves the modified file. Includes checks for missing files, undefined names, and malformed formulas.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook that contains the named range (e.g., "PublishRange")
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range definition
                // Assumes the named range is already defined in the workbook
                Name publishRangeName = workbook.Worksheets.Names["PublishRange"];
                if (publishRangeName == null)
                {
                    Console.WriteLine("Named range 'PublishRange' not found.");
                    return;
                }

                // The RefersTo property returns a formula like "=Sheet1!$A$1:$B$10"
                string refersTo = publishRangeName.RefersTo;
                if (refersTo.StartsWith("="))
                    refersTo = refersTo.Substring(1); // remove leading '='

                // Split sheet name and address
                int exclPos = refersTo.IndexOf('!');
                if (exclPos < 0)
                {
                    Console.WriteLine("Invalid RefersTo format for named range.");
                    return;
                }

                string sheetName = refersTo.Substring(0, exclPos);
                string rangeAddress = refersTo.Substring(exclPos + 1);

                // Get the worksheet that contains the range
                Worksheet ws = workbook.Worksheets[sheetName];
                if (ws == null)
                {
                    Console.WriteLine($"Worksheet \"{sheetName}\" not found.");
                    return;
                }

                // Create a Range object for the address (use fully qualified name to avoid ambiguity)
                Aspose.Cells.Range range = ws.Cells.CreateRange(rangeAddress);

                // Iterate through all cells in the range and replace "Draft" with "Final"
                int firstRow = range.FirstRow;
                int firstColumn = range.FirstColumn;
                int rowCount = range.RowCount;
                int columnCount = range.ColumnCount;

                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        Cell cell = ws.Cells[firstRow + i, firstColumn + j];
                        if (cell.Type == CellValueType.IsString)
                        {
                            string text = cell.StringValue;
                            if (text.Contains("Draft"))
                            {
                                // Replace the word "Draft" with "Final"
                                cell.PutValue(text.Replace("Draft", "Final"));
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
