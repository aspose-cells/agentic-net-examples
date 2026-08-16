// Title: Find and Highlight Dates Older Than One Year in a Named Range with Aspose.Cells (C#)
// Description: Loads an Excel workbook, accesses the named range "MyDateRange", defines its CellArea, and scans each cell for DateTime values earlier than today minus one year. Matching cells are printed, highlighted with a LightSalmon background, and the workbook is saved as Output.xlsx.
// Keywords: Aspose.Cells | C# | FindOptions | named range | date comparison | highlight cells | CellArea | Excel automation | old dates | style formatting
// Common Searches: Aspose.Cells find dates older than a year | C# highlight old dates in Excel named range | Use FindOptions to locate past dates in Aspose.Cells | filter dates before a cutoff date with Aspose.Cells | apply style to cells based on date value in C#
// Developer Intent: Search a specific named range for DateTime cells older than one year and apply visual highlighting.
// Use Cases: Flag expired contract dates in financial reports before distribution. | Create an audit list of all timestamps older than a year from a predefined range. | Automatically color‑code outdated entries in spreadsheets for compliance checks.
// AI Prompts: Generate C# code that uses Aspose.Cells FindOptions to search a named range for DateTime cells older than one year and apply a custom highlight style. | Show how to retrieve a named range, define a CellArea, and use FindOptions with a date predicate to locate and format old dates in an Excel file.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsFindOldDates
{
    // Loads an Excel workbook, accesses the named range "MyDateRange", defines its CellArea, and scans each cell for DateTime values earlier than today minus one year. Matching cells are printed, highlighted with a LightSalmon background, and the workbook is saved as Output.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "Input.xlsx";
                const string outputPath = "Output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range "MyDateRange"
                AsposeRange namedRange = workbook.Worksheets.GetRangeByName("MyDateRange");
                if (namedRange == null)
                {
                    Console.WriteLine("Named range 'MyDateRange' not found.");
                    return;
                }

                // Define the search area based on the named range
                CellArea rangeArea = new CellArea
                {
                    StartRow = namedRange.FirstRow,
                    StartColumn = namedRange.FirstColumn,
                    EndRow = namedRange.FirstRow + namedRange.RowCount - 1,
                    EndColumn = namedRange.FirstColumn + namedRange.ColumnCount - 1
                };

                // Cutoff date: one year ago from today
                DateTime cutoffDate = DateTime.Now.AddYears(-1);

                Console.WriteLine("Cells containing dates older than one year:");

                // Scan cells and output those older than the cutoff date
                for (int row = rangeArea.StartRow; row <= rangeArea.EndRow; row++)
                {
                    for (int col = rangeArea.StartColumn; col <= rangeArea.EndColumn; col++)
                    {
                        Cell cell = namedRange.Worksheet.Cells[row, col];
                        if (cell.Type == CellValueType.IsDateTime && cell.DateTimeValue < cutoffDate)
                        {
                            Console.WriteLine($"{cell.Name} = {cell.DateTimeValue:d}");
                        }
                    }
                }

                // Highlight the found cells
                Style highlightStyle = workbook.CreateStyle();
                highlightStyle.ForegroundColor = System.Drawing.Color.LightSalmon;
                highlightStyle.Pattern = BackgroundType.Solid;

                for (int row = rangeArea.StartRow; row <= rangeArea.EndRow; row++)
                {
                    for (int col = rangeArea.StartColumn; col <= rangeArea.EndColumn; col++)
                    {
                        Cell cell = namedRange.Worksheet.Cells[row, col];
                        if (cell.Type == CellValueType.IsDateTime && cell.DateTimeValue < cutoffDate)
                        {
                            cell.SetStyle(highlightStyle);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
