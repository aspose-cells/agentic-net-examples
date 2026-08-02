// Title: Find numeric values >1000 in the "Budget" named range with FindOptions – Aspose.Cells C#
// Description: Creates a workbook, defines a named range "Budget" (B2:B5), configures FindOptions with a CellArea that matches the range, and extracts cells whose numeric values exceed 1000, then saves the file.
// Keywords: Aspose.Cells FindOptions | named range search | Budget range | values greater than 1000 | C# Excel cell filtering | CellArea SetRange example
// Common Searches: Aspose.Cells FindOptions named range example | search cells >1000 in Excel using C# | how to filter numeric values in a named range Aspose.Cells | set range for FindOptions Aspose.Cells .NET
// Developer Intent: Retrieve all numeric cells inside the named range "Budget" whose values are higher than 1000.
// Use Cases: Highlight expense items that exceed a budget limit. | Validate financial worksheets before approval. | Generate reports of high‑value transactions.
// AI Prompts: Show C# code using Aspose.Cells FindOptions to return cells with values >1000 in the named range "Budget". | Explain how to combine FindOptions.SetRange and a CellArea to filter numeric cells in a workbook. | Provide an Aspose.Cells example that searches a named range for values above a threshold without iterating manually.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFindOptionsDemo
{
    // Creates a workbook, defines a named range "Budget" (B2:B5), configures FindOptions with a CellArea that matches the range, and extracts cells whose numeric values exceed 1000, then saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (some values above and below 1000)
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("Office Supplies");
                sheet.Cells["B2"].PutValue(750);
                sheet.Cells["A3"].PutValue("Equipment");
                sheet.Cells["B3"].PutValue(1250);
                sheet.Cells["A4"].PutValue("Travel");
                sheet.Cells["B4"].PutValue(2000);
                sheet.Cells["A5"].PutValue("Utilities");
                sheet.Cells["B5"].PutValue(500);

                // Define a named range "Budget" that covers the Amount column (B2:B5)
                Aspose.Cells.Range budgetRange = sheet.Cells.CreateRange("B2", "B5");
                budgetRange.Name = "Budget";

                // Retrieve the named range using the workbook's worksheet collection
                Aspose.Cells.Range namedRange = workbook.Worksheets.GetRangeByName("Budget");
                if (namedRange == null)
                {
                    Console.WriteLine("Named range 'Budget' not found.");
                    return;
                }

                // Configure FindOptions to limit the search to the named range
                FindOptions findOptions = new FindOptions
                {
                    LookInType = LookInType.Values,
                    // Default LookAtType is Whole, so we omit explicit setting to avoid version issues
                    SearchOrderByRows = true
                };

                // Convert the Range to a CellArea and assign it to FindOptions
                CellArea area = new CellArea
                {
                    StartRow = namedRange.FirstRow,
                    StartColumn = namedRange.FirstColumn,
                    EndRow = namedRange.FirstRow + namedRange.RowCount - 1,
                    EndColumn = namedRange.FirstColumn + namedRange.ColumnCount - 1
                };
                findOptions.SetRange(area);

                // Iterate through each cell in the named range and output values > 1000
                Console.WriteLine("Numeric values greater than 1000 in named range 'Budget':");
                for (int row = area.StartRow; row <= area.EndRow; row++)
                {
                    Cell cell = sheet.Cells[row, area.StartColumn];
                    if (cell.Type == CellValueType.IsNumeric && cell.DoubleValue > 1000)
                    {
                        Console.WriteLine($"Cell {cell.Name}: {cell.DoubleValue}");
                    }
                }

                // Save the workbook
                string outputPath = "FindOptionsBudgetDemo.xlsx";
                try
                {
                    // Ensure the directory exists
                    string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
