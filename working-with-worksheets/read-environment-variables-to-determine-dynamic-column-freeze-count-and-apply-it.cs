// Title: How to freeze a variable number of columns in an Excel sheet with Aspose.Cells for .NET using an environment variable
// AI Prompts: Generate C# code that reads a FREEZE_COLUMNS environment variable, validates it as a positive integer, and uses Aspose.Cells FreezePanes to freeze that many columns while leaving rows unfrozen. | Create a C# example that defaults to no column freeze when the environment variable is missing, empty, or not a positive integer, then saves the workbook as output.xlsx. | Extend the snippet to also read a FREEZE_ROWS environment variable and apply both row and column freezing with Aspose.Cells FreezePanes.
// Common Searches: aspnet read environment variable to set column freeze in Aspose.Cells workbook | c# Aspose.Cells FreezePanes column count from config value | how to dynamically freeze columns in Excel using Aspose.Cells and env variable | set freeze panes based on environment variable in .NET Excel generation
// Tags: Aspose.Cells FreezePanes column freeze | environment variable driven Excel pane freezing .NET | dynamic column freeze in Aspose.Cells | C# FreezePanes based on config value | programmatic Excel pane freezing with Aspose.Cells

using System;
using Aspose.Cells;

// The program reads the FREEZE_COLUMNS environment variable, parses it as a positive integer, creates a Workbook, accesses the first worksheet, and if the value is greater than zero calls FreezePanes(0, count, 0, count) to freeze that many columns (no rows frozen). The workbook is saved as output.xlsx, and missing or invalid values result in no frozen panes.
class Program
{
    static void Main()
    {
        try
        {
            // Read the environment variable that specifies how many columns to freeze.
            // If the variable is missing or invalid, default to 0 (no freeze).
            int freezeColumnCount = 0;
            string envValue = Environment.GetEnvironmentVariable("FREEZE_COLUMNS");
            if (!string.IsNullOrEmpty(envValue) && int.TryParse(envValue, out int parsed) && parsed > 0)
            {
                freezeColumnCount = parsed;
            }

            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Apply column freeze based on the environment variable.
            // Freeze rows = 0 (no row freeze), columns = freezeColumnCount.
            if (freezeColumnCount > 0)
            {
                // FreezePanes(row, column, totalRows, totalColumns)
                // row = 0 (no rows frozen), column = freezeColumnCount (first scrollable column),
                // totalRows = 0, totalColumns = freezeColumnCount (freeze the specified columns).
                sheet.FreezePanes(0, freezeColumnCount, 0, freezeColumnCount);
            }

            // Save the workbook.
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
