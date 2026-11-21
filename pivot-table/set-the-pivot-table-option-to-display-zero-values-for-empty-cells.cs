// Title: Configure an Aspose.Cells PivotTable in C# to display zero for blank and error cells
// Common Searches: Aspose.Cells C# set pivot table to display 0 for null values | Show zero instead of blank cells in an Aspose.Cells pivot table using .NET | Replace error values with 0 in an Aspose.Cells PivotTable programmatically | Enable custom null and error strings for pivot tables in C# Aspose.Cells | Aspose.Cells pivot table display zero for empty cells example
// Tags: Aspose.Cells pivot table null string setting | pivot table display zero for blank cells C# | custom error string configuration Aspose.Cells | C# workbook save after pivot table changes | Aspose.Cells pivot table empty cell handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

// The sample loads an existing Excel workbook, verifies a pivot table exists on the first worksheet, and configures the pivot table to show the string "0" for both null and error values by enabling DisplayNullString and DisplayErrorString and setting NullString and ErrorString to "0". The modified workbook is then saved to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Access the first pivot table on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Enable display of custom strings for null and error values
            pivotTable.DisplayNullString = true;
            pivotTable.NullString = "0";

            // Enable display of custom strings for error values
            pivotTable.DisplayErrorString = true;
            pivotTable.ErrorString = "0";

            // Save the workbook with the updated pivot table settings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display the error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
