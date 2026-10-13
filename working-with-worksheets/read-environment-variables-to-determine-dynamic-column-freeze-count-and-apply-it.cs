// Title: Freeze a configurable number of columns in an Excel worksheet using Aspose.Cells for .NET with an environment variable
// AI Prompts: Generate C# code that reads the FREEZE_COLUMN_COUNT environment variable and uses Worksheet.FreezePanes to freeze that many columns in the first sheet of an Excel file with Aspose.Cells. | Show how to handle missing or invalid environment values when applying column freeze panes in a .NET console app that uses Aspose.Cells. | Provide a complete example that loads input.xlsx, freezes a variable number of columns based on a system variable, and saves the result as output.xlsx using Aspose.Cells.
// Common Searches: asp.net read environment variable to set freeze panes in Excel with Aspose.Cells | c# freeze first N columns in workbook using Aspose.Cells where N is from env variable | how to programmatically freeze columns based on configuration in Aspose.Cells | using Worksheet.FreezePanes with dynamic column count in .NET
// Tags: Worksheet.FreezePanes variable column index | Aspose.Cells system variable configuration | Excel column freeze automation .NET | Workbook load save with configurable freeze panes | handle invalid FREEZE_COLUMN_COUNT value

using System;
using System.IO;
using Aspose.Cells;

// The program reads the FREEZE_COLUMN_COUNT environment variable, loads input.xlsx, freezes the specified number of columns in the first worksheet using Worksheet.FreezePanes, and saves the modified workbook as output.xlsx, including error handling for missing files and invalid variable values.
class Program
{
    static void Main()
    {
        try
        {
            // Read the environment variable that specifies how many columns to freeze
            string envValue = Environment.GetEnvironmentVariable("FREEZE_COLUMN_COUNT");
            int freezeColumnCount = 0;
            if (!int.TryParse(envValue, out freezeColumnCount))
            {
                // Default to 0 (no freeze) if the variable is missing or invalid
                freezeColumnCount = 0;
            }

            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Apply freeze panes:
            //   row = 0 (no rows frozen)
            //   column = freezeColumnCount (split position)
            //   totalRows = 0 (no rows frozen)
            //   totalColumns = freezeColumnCount (freeze the specified columns)
            sheet.FreezePanes(0, freezeColumnCount, 0, freezeColumnCount);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
