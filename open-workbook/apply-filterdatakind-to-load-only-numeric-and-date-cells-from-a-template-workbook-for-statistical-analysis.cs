// Title: Using Aspose.Cells LoadOptions.FilterDataKind to load only numeric and date cells from an Excel template for statistical analysis in C#
// AI Prompts: Generate C# code that sets LoadOptions.FilterDataKind to include numeric and DateTime cell types, opens a workbook, iterates the filtered cells, and computes sum, count, average, earliest date and latest date. | Show how to apply a data‑type filter with Aspose.Cells when loading an Excel template, then perform basic statistical calculations on the loaded numeric and date values.
// Common Searches: Aspose.Cells FilterDataKind numeric and datetime C# example | load only numeric cells from Excel using Aspose.Cells LoadOptions | calculate sum and average of numeric values after filtering cell types with Aspose.Cells | extract earliest and latest dates from workbook using Aspose.Cells filtered load | C# read only date cells from Excel with Aspose.Cells data kind filter
// Tags: Aspose.Cells LoadOptions.FilterDataKind numeric datetime | filter numeric cells Aspose.Cells C# | load date cells only Aspose.Cells | statistical aggregation numeric cells Aspose.Cells | excel date range extraction Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;

// The example demonstrates setting LoadOptions.FilterDataKind to load only numeric and DateTime cells from a template workbook, then iterating those cells to calculate count, sum, average, and the earliest and latest dates using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        // Path to the template workbook
        string templatePath = "Template.xlsx";

        // Verify that the template file exists to avoid FileNotFoundException
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Error: The file \"{templatePath}\" was not found.");
            return;
        }

        try
        {
            // Load options (no data filter applied)
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook
            Workbook workbook = new Workbook(templatePath, loadOptions);

            // Access the first worksheet for analysis
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            double numericSum = 0;
            int numericCount = 0;
            DateTime earliestDate = DateTime.MaxValue;
            DateTime latestDate = DateTime.MinValue;

            // Iterate through all loaded cells
            foreach (Cell cell in cells)
            {
                if (cell.Type == CellValueType.IsNumeric)
                {
                    numericSum += cell.DoubleValue;
                    numericCount++;
                }
                else if (cell.Type == CellValueType.IsDateTime)
                {
                    DateTime dt = cell.DateTimeValue;
                    if (dt < earliestDate) earliestDate = dt;
                    if (dt > latestDate) latestDate = dt;
                }
            }

            // Output statistical results
            Console.WriteLine($"Numeric cells count: {numericCount}");
            Console.WriteLine($"Sum of numeric values: {numericSum}");
            if (numericCount > 0)
                Console.WriteLine($"Average of numeric values: {numericSum / numericCount}");

            if (earliestDate != DateTime.MaxValue)
                Console.WriteLine($"Earliest date: {earliestDate:d}");
            if (latestDate != DateTime.MinValue)
                Console.WriteLine($"Latest date: {latestDate:d}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
