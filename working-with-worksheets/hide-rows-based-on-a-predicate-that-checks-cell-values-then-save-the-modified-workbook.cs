// Title: Hide rows with negative values in column A using Aspose.Cells for .NET and save the workbook
// AI Prompts: Write C# code with Aspose.Cells that loads an Excel file, hides every row where column A contains a value less than zero, and saves the result to a new file. | Create a .NET method that iterates through a worksheet's rows, sets the IsHidden property on rows meeting a numeric predicate, and persists the workbook using Aspose.Cells.
// Common Searches: aspnet hide rows in excel where column value is negative using aspose.cells | c# aspose.cells hide rows based on cell condition and save workbook | how to programmatically hide rows with negative numbers in column A with Aspose.Cells | Aspose.Cells hide rows with predicate and export to new Excel file | filter out rows by value and hide them in Excel using Aspose.Cells .NET
// Tags: hide rows based on cell value Aspose.Cells | set IsHidden property C# Excel | filter rows by numeric predicate Aspose.Cells | save modified workbook Aspose.Cells .NET | use MaxDataRow to determine last row Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (or creates a new one), iterates through all rows up to the last populated row, checks the numeric value in column A, hides rows where the value is less than zero by setting the IsHidden property, ensures the output directory exists, and saves the modified workbook to a specified path while handling potential errors.
class HideRowsExample
{
    static void Main()
    {
        try
        {
            // Paths for input and output workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Optional: add sample data to the default sheet
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue(10);
                ws.Cells["A2"].PutValue(-5);
                ws.Cells["A3"].PutValue(15);
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Find the last row that contains data
            int lastDataRow = worksheet.Cells.MaxDataRow;

            // Hide rows where the value in column A is less than 0
            for (int rowIndex = 0; rowIndex <= lastDataRow; rowIndex++)
            {
                Cell cell = worksheet.Cells[rowIndex, 0];
                if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double numericValue))
                {
                    if (numericValue < 0)
                    {
                        // Use IsHidden property to hide the row
                        worksheet.Cells.Rows[rowIndex].IsHidden = true;
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
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
