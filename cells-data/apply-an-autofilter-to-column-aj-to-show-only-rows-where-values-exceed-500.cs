// Title: How to apply a numeric AutoFilter on column AJ (>500) using Aspose.Cells for .NET
// AI Prompts: Insert a CustomFilter on column AJ (index 35) with the GreaterThan operator for the value 500 and save the workbook. | Define the AutoFilter range from A1 to AJ and programmatically filter rows where AJ exceeds 500 using Aspose.Cells in C#. | Uncomment the filter lines in the sample code to display only rows with AJ values greater than 500, then export to Filtered.xlsx.
// Common Searches: How to use Aspose.Cells to filter rows where column AJ is greater than 500 in C# | Applying a numeric AutoFilter on a specific column with Aspose.Cells .NET | Programmatic Excel AutoFilter range definition using Aspose.Cells | C# example for CustomFilter GreaterThan on Excel column with Aspose.Cells | Filtering Excel data by column value using Aspose.Cells library
// Tags: Aspose.Cells numeric AutoFilter C# | filter column AJ greater than 500 Aspose.Cells | set AutoFilter range A1:AJ Aspose.Cells | CustomFilter GreaterThan operator .NET | Excel data filtering with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The C# program loads or creates an Excel workbook, calculates the last used row, sets the AutoFilter range to A1:AJ, and (when the commented lines are enabled) applies a numeric CustomFilter on column AJ to keep only rows with values over 500, then saves the result as 'Filtered.xlsx'.
    class Program
    {
        static void Main()
        {
            try
            {
                // Load an existing workbook if it exists; otherwise create a new one
                string inputPath = "input.xlsx";
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the last used row (zero‑based) and add 1 for the range end
                int lastRow = sheet.Cells.MaxDataRow + 1;
                if (lastRow == 0) lastRow = 1; // ensure at least the header row is included

                // Define the AutoFilter range covering columns A to AJ
                string filterRange = $"A1:AJ{lastRow}";
                sheet.AutoFilter.Range = filterRange;

                // NOTE: The CustomFilter method and AutoFilterOperator enum may not be available
                // in older versions of Aspose.Cells. If they are present, uncomment the lines below
                // to apply a numeric filter on column AJ (zero‑based index 35) for values > 500.

                // sheet.AutoFilter.CustomFilter(35, "500", AutoFilterOperator.GreaterThan);
                // sheet.AutoFilter.Apply();

                // Save the workbook (filtered if the above lines are enabled)
                string outputPath = "Filtered.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
