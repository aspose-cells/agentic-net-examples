// Title: Hide rows in an Excel worksheet where column A is negative using Aspose.Cells for .NET and save the file
// Common Searches: C# Aspose.Cells hide rows where column A value is less than zero | How to programmatically hide Excel rows based on a cell predicate with Aspose.Cells .NET | Iterate through rows and set Row.IsHidden in Aspose.Cells C# example | Save an Excel workbook after conditionally hiding rows using Aspose.Cells | Filter rows by numeric value in column A and hide them with Aspose.Cells for .NET
// Tags: conditional row hiding Aspose.Cells .NET | Row.IsHidden property C# Excel | hide rows based on cell value Aspose.Cells | save workbook after row visibility changes Aspose.Cells | numeric predicate row filtering Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an input.xlsx workbook, iterates over all data rows on the first worksheet, checks the numeric value in column A, hides any row whose value is less than zero by setting Row.IsHidden, and then saves the updated workbook to output.xlsx while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the range of rows that contain data
            int firstDataRow = sheet.Cells.MinRow;
            int lastDataRow = sheet.Cells.MaxDataRow;

            // Iterate through each row and hide it if the predicate is satisfied
            // Example predicate: hide rows where the value in column A is less than 0
            for (int rowIndex = firstDataRow; rowIndex <= lastDataRow; rowIndex++)
            {
                Cell cell = sheet.Cells[rowIndex, 0]; // Column A (index 0)

                // Check if the cell contains a numeric value and apply the predicate
                if (cell.Value != null && double.TryParse(cell.StringValue, out double numericValue))
                {
                    if (numericValue < 0)
                    {
                        // Hide the entire row using the Row collection
                        Row row = sheet.Cells.Rows[rowIndex];
                        row.IsHidden = true;
                    }
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
