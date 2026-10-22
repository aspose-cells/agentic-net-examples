// Title: Hide rows with "Inactive" status after importing a CSV file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# hide rows where column value equals Inactive after loading CSV | How to filter out rows with status Inactive in a workbook imported from CSV using Aspose.Cells | C# Aspose.Cells hide rows based on cell text after CSV conversion | Set row IsHidden property for rows with "Inactive" status in Aspose.Cells .NET | Programmatically hide rows in Excel after converting CSV with Aspose.Cells
// Tags: hide rows based on cell value Aspose.Cells | filter rows after CSV import Aspose.Cells | set row IsHidden property C# | load CSV into workbook Aspose.Cells | conditional row visibility Excel .NET | status column filtering Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a CSV file into an Aspose.Cells workbook, iterates through the data rows, hides any row whose "Status" column (column B) contains the value "Inactive", and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        const string inputFile = "input.csv";
        const string outputFile = "output.xlsx";

        // Verify that the input CSV file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file not found: {Path.GetFullPath(inputFile)}");
            return;
        }

        try
        {
            // Load CSV data into a new workbook using appropriate LoadOptions
            var loadOptions = new LoadOptions(LoadFormat.Csv);
            var workbook = new Workbook(inputFile, loadOptions);

            // Reference to the first worksheet (contains the imported CSV data)
            var sheet = workbook.Worksheets[0];

            // Assume the "Status" column is column B (zero‑based index = 1)
            const int statusColumnIndex = 1;

            // Determine the last row with data
            int lastRow = sheet.Cells.MaxDataRow;

            // Loop through each data row (skip header row at index 0)
            for (int row = 1; row <= lastRow; row++)
            {
                // Check if the cell in the Status column equals "Inactive"
                var cellValue = sheet.Cells[row, statusColumnIndex].StringValue;
                if (cellValue.Equals("Inactive", StringComparison.OrdinalIgnoreCase))
                {
                    // Hide the entire row
                    sheet.Cells.Rows[row].IsHidden = true;
                }
            }

            // Ensure the output directory exists
            var outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook to a new Excel file
            workbook.Save(outputFile);
            Console.WriteLine($"Processing complete. Output saved to {Path.GetFullPath(outputFile)}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
