// Title: C# Aspose.Cells example: Highlight numeric values > 1000 in the named range “Budget” using FindOptions
// Description: Loads an Excel workbook, gets the named range "Budget", defines a CellArea, configures FindOptions to search only that area, finds cells with numeric values greater than 1000, applies a yellow background style, reports the matches, and saves the updated file.
// Keywords: Aspose.Cells | FindOptions | C# | .NET | named range | Budget range | numeric search | values > 1000 | highlight cells | CellArea | SetRange | Excel automation | financial spreadsheet
// Common Searches: Aspose.Cells FindOptions example for named range | C# highlight cells greater than 1000 in Excel | How to limit FindOptions to a specific range | Search numeric values in a named range with Aspose.Cells | Apply style to cells that exceed a threshold using Aspose
// Developer Intent: Locate every numeric cell whose value exceeds 1000 inside the "Budget" named range and visually mark those cells.
// Use Cases: Spot budget overruns by automatically coloring high‑value entries. | Extract large expense items for a financial summary report. | Provide visual cues in budgeting worksheets for quick review.
// AI Prompts: Generate C# code that uses Aspose.Cells FindOptions to search a named range for numbers above a threshold and apply a custom style. | Show how to set a CellArea with SetRange so FindOptions scans only the "Budget" range. | Create an example that iterates over cells in a named range, collects values > 1000, highlights them, and saves the workbook.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsFindOptionsExample
{
    // Loads an Excel workbook, gets the named range "Budget", defines a CellArea, configures FindOptions to search only that area, finds cells with numeric values greater than 1000, applies a yellow background style, reports the matches, and saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "InputWorkbook.xlsx";
                const string outputPath = "OutputWorkbook.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook that contains a named range called "Budget"
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range "Budget"
                AsposeRange budgetRange = workbook.Worksheets.GetRangeByName("Budget");
                if (budgetRange == null)
                {
                    Console.WriteLine("Named range 'Budget' not found.");
                    return;
                }

                // Convert the Range to a CellArea for FindOptions
                CellArea searchArea = new CellArea
                {
                    StartRow = budgetRange.FirstRow,
                    StartColumn = budgetRange.FirstColumn,
                    EndRow = budgetRange.FirstRow + budgetRange.RowCount - 1,
                    EndColumn = budgetRange.FirstColumn + budgetRange.ColumnCount - 1
                };

                // Configure FindOptions to limit the search to the "Budget" area
                FindOptions findOptions = new FindOptions
                {
                    LookInType = LookInType.Values // Search only cell values
                    // LookAtType defaults to Whole, which is suitable for numeric comparison
                };
                findOptions.SetRange(searchArea);

                // Prepare a style to highlight cells that satisfy the condition (> 1000)
                Style highlightStyle = workbook.CreateStyle();
                highlightStyle.ForegroundColor = System.Drawing.Color.Yellow;
                highlightStyle.Pattern = BackgroundType.Solid;

                // Collect cells with values > 1000 and apply the highlight style
                List<Cell> cellsGreaterThanThousand = new List<Cell>();
                Worksheet sheet = budgetRange.Worksheet;
                for (int row = searchArea.StartRow; row <= searchArea.EndRow; row++)
                {
                    for (int col = searchArea.StartColumn; col <= searchArea.EndColumn; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell.Type == CellValueType.IsNumeric && cell.DoubleValue > 1000)
                        {
                            cellsGreaterThanThousand.Add(cell);
                            cell.SetStyle(highlightStyle);
                        }
                    }
                }

                // Output the results
                Console.WriteLine($"Found {cellsGreaterThanThousand.Count} cells with values > 1000 in the 'Budget' range.");
                foreach (Cell c in cellsGreaterThanThousand)
                {
                    Console.WriteLine($"Cell {c.Name}: {c.DoubleValue}");
                }

                // Save the workbook with the highlighted cells
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
