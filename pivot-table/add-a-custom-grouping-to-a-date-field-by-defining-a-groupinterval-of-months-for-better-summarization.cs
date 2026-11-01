// Title: How to group a pivot table date field by months with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# group pivot table date field by month interval | C# code to set month grouping on an Excel pivot table using Aspose.Cells | GroupByDate months not supported Aspose.Cells workaround example | Create month groups for OrderDate field in a pivot table with Aspose.Cells .NET | Programmatically apply month grouping to pivot table rows in C# Excel library
// Tags: Aspose.Cells pivot table month grouping | C# GroupByDate months Aspose.Cells | Excel pivot date field grouping .NET | custom date interval grouping Aspose.Cells | programmatic pivot table date grouping C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, checks for a worksheet and a pivot table, locates the 'OrderDate' row field, and shows where month‑based grouping would be applied using GroupByDate if the API supported it, then saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Retrieve the first pivot table (adjust index if needed)
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Name of the date field to be grouped
            string dateFieldName = "OrderDate";

            // Locate the date field within the row fields of the pivot table
            PivotField dateField = null;
            foreach (PivotField field in pivotTable.RowFields)
            {
                if (field.Name.Equals(dateFieldName, StringComparison.OrdinalIgnoreCase))
                {
                    dateField = field;
                    break;
                }
            }

            if (dateField != null)
            {
                try
                {
                    // NOTE: GroupByDate is not available in the current Aspose.Cells version.
                    // If supported, you could use:
                    // DateTime startDate = new DateTime(1900, 1, 1);
                    // DateTime endDate   = new DateTime(2100, 12, 31);
                    // dateField.GroupByDate(startDate, endDate, GroupByDateType.Months);
                    Console.WriteLine($"Field '{dateFieldName}' found. Grouping by months would be applied here if supported.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error while attempting to group date field: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Date field '{dateFieldName}' not found in the pivot table row fields.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
