// Title: Hide rows with negative numbers in column A using Aspose.Cells for .NET and save the workbook
// AI Prompts: Load an Excel file with Aspose.Cells, hide every row where column A contains a value less than zero, and write the result to a new file. | Iterate through all used rows in a worksheet, evaluate a numeric predicate on a specific cell, set the row's IsHidden property, and save the workbook. | Create a workbook if it doesn't exist, apply a custom row‑hiding rule based on cell values, and export the modified workbook using Aspose.Cells for C#.
// Common Searches: Aspose.Cells hide rows where column A value is negative C# | C# hide Excel rows based on cell value using Aspose.Cells | How to programmatically hide rows in an Excel worksheet with Aspose.Cells .NET | Save workbook after hiding rows with Aspose.Cells in C# | Iterate rows and set IsHidden property in Aspose.Cells example
// Tags: hide rows based on numeric condition Aspose.Cells | Row.IsHidden property C# | filter worksheet rows by cell value Aspose.Cells | save modified workbook Aspose.Cells .NET | load or create workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program loads an existing Excel file (or creates a new one), scans each used row in the first worksheet, hides any row where the value in column A is negative by setting the row's IsHidden property, and then saves the updated workbook to output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Work with the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Column index to evaluate (0 = column A)
            int columnIndex = 0;

            // Determine the last row that contains data
            int lastRow = cells.MaxDataRow;

            // Iterate through each used row
            for (int row = 0; row <= lastRow; row++)
            {
                Cell cell = cells[row, columnIndex];

                // Hide the row if the cell contains a numeric value less than 0
                if (cell.Type == CellValueType.IsNumeric && cell.DoubleValue < 0)
                {
                    // Use Row.IsHidden property to hide the row
                    worksheet.Cells.Rows[row].IsHidden = true;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
