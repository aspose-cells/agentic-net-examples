// Title: Hide non‑Completed rows in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that iterates through a worksheet and hides every row whose 'Status' column value is not "Completed", then saves the workbook. | Create a reusable C# method that takes a Worksheet, a column index, and a target string, and uses Aspose.Cells to set IsHidden = true for rows where the cell value differs from the target. | Adapt the example to hide rows based on a different column name (e.g., "Phase") and a custom keyword, ensuring the hidden rows are persisted in a new .xlsx file with Aspose.Cells.
// Common Searches: Aspose.Cells C# hide rows where column value is not Completed | filter Excel rows by status column using Aspose.Cells .NET | C# program to hide rows that don't match a specific text in an .xlsx file | how to set row visibility based on cell content with Aspose.Cells for .NET
// Tags: Aspose.Cells hide rows by cell value | C# Excel conditional row visibility | Aspose.Cells row hiding based on column criteria | Excel status column filter using Aspose.Cells | programmatic row hiding .NET Excel

using System;
using System.IO;
using Aspose.Cells;

// The C# example loads an XLSX workbook with Aspose.Cells, identifies the 'Status' column (index 2), iterates over all data rows, and sets each row's IsHidden property to true when the cell value is not "Completed". The modified workbook is then saved to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Index of the "Status" column (0‑based). Adjust if the column is different.
            int statusColumnIndex = 2; // e.g., column C

            // Determine the first and last rows that contain data.
            // Assuming row 0 holds headers, data starts at row 1.
            int firstDataRow = 1;
            int lastDataRow = sheet.Cells.MaxDataRow;

            // Hide rows where the Status column value is not "Completed"
            for (int row = firstDataRow; row <= lastDataRow; row++)
            {
                Cell statusCell = sheet.Cells[row, statusColumnIndex];
                bool isCompleted = string.Equals(
                    statusCell.StringValue,
                    "Completed",
                    StringComparison.OrdinalIgnoreCase);

                // Use Cells.Rows to hide the row
                sheet.Cells.Rows[row].IsHidden = !isCompleted;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
