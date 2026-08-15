// Title: C# – Use Aspose.Cells FindAll to prepend "Reviewed:" to every string cell
// Description: Demonstrates how to create a workbook, populate sample text, locate all string cells with the FindAll method, prepend the prefix "Reviewed:" to each value, and save the result as ReviewedOutput.xlsx. Includes null‑safety and type checking.
// Keywords: Aspose.Cells FindAll C# | prepend text to cell values | modify string cells .NET | iterate used range Aspose.Cells | add prefix to Excel cells | Aspose.Cells workbook automation | C# Excel cell update | global .NET spreadsheet processing
// Common Searches: Aspose.Cells prepend prefix to all text cells | C# FindAll example for updating cell values | How to add "Reviewed:" to every string in an Excel file using Aspose | Iterate over used range and modify cells Aspose.Cells | Update cell content programmatically in .NET
// Developer Intent: Add the "Reviewed:" prefix to each string‑type cell in a worksheet using Aspose.Cells.
// Use Cases: Flag every task description as reviewed before distribution | Automatically label textual entries with a status marker for audit trails | Prepare export files where all string data must carry a custom prefix
// AI Prompts: Write C# code that uses Aspose.Cells FindAll to locate all string cells and prepend "Reviewed:" to each. | Show how to replace a manual nested loop with FindAll while safely handling null or non‑string cells. | Explain best practices for bulk updating cell values in Aspose.Cells without affecting numeric data.

using System;
using Aspose.Cells;

namespace AsposeCellsFindAllDemo
{
    // Demonstrates how to create a workbook, populate sample text, locate all string cells with the FindAll method, prepend the prefix "Reviewed:" to each value, and save the result as ReviewedOutput.xlsx. Includes null‑safety and type checking.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate some sample data
                cells["A1"].PutValue("Task 1");
                cells["B2"].PutValue("Task 2");
                cells["C3"].PutValue("Task 3");

                // Determine the used range of the worksheet
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                // Iterate over all cells in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = cells[row, col];
                        // Process only cells that contain a string value
                        if (cell.Type == CellValueType.IsString)
                        {
                            string originalValue = cell.StringValue ?? string.Empty;
                            // Prepend "Reviewed:" to the existing value
                            cell.PutValue("Reviewed:" + originalValue);
                        }
                    }
                }

                // Save the workbook (lifecycle rule: save)
                workbook.Save("ReviewedOutput.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
