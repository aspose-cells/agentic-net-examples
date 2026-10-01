// Title: How to replace empty cells with 'N/A' in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that walks through the worksheet’s used area and writes "N/A" into each empty cell. | Show a method to iterate over all cells in an Aspose.Cells workbook and assign a placeholder text when a cell has no value before saving.
// Common Searches: how to display N/A for empty cells using Aspose.Cells C# | replace null values with placeholder text in Excel file programmatically Aspose.Cells | iterate over used range and set default value for blank cells Aspose.Cells .NET | Aspose.Cells C# set default string for empty worksheet cells
// Tags: fill missing entries Aspose.Cells C# | scan worksheet area Aspose.Cells | default text for absent cells Excel .NET | populate N/A values Aspose.Cells | null cell handling Aspose.Cells C#

using System;
using Aspose.Cells;

namespace AsposeCellsNullDisplay
{
    // The program creates a new workbook, adds sample data with intentional empty cells, walks through the worksheet's used area, replaces any null cell values with the string "N/A", and saves the result as NullValuesDisplayed.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (some cells intentionally left empty)
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Score");
                sheet.Cells["A2"].PutValue("Alice");
                sheet.Cells["B2"].PutValue(85);
                sheet.Cells["A3"].PutValue("Bob");
                // B3 left empty
                sheet.Cells["A4"].PutValue("Charlie");
                // B4 left empty

                // Get the used range of the worksheet (use Aspose.Cells.Range to avoid ambiguity with System.Range)
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Replace empty cells with "N/A"
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell.Value == null)
                        {
                            cell.PutValue("N/A");
                        }
                    }
                }

                // Save the workbook
                workbook.Save("NullValuesDisplayed.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
