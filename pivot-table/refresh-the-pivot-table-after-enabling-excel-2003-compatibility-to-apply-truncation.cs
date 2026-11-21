// Title: Refresh all pivot tables after enabling Excel 2003 compatibility with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set Excel 2003 compatibility and refresh pivot tables | How to programmatically refresh pivot tables after changing CompatibilityOptions in Aspose.Cells | Enable Excel2003 mode and update pivot cache using Aspose.Cells .NET | Set RefreshDataOnOpen for pivot tables with Aspose.Cells C# example | Refresh pivot tables in a workbook after applying Excel 2003 compatibility with Aspose.Cells
// Tags: Excel2003 compatibility option Aspose.Cells | pivot table data refresh Aspose.Cells | RefreshData method Aspose.Cells | RefreshDataOnOpen property Aspose.Cells | programmatic pivot cache update .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The sample loads a workbook, uses reflection to set the CompatibilityOptions to Excel2003 when available, iterates through each worksheet and its pivot tables to call RefreshData, optionally enables RefreshDataOnOpen, and saves the updated file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Set Excel 2003 compatibility if the property exists
                var settings = workbook.Settings;
                var compatProp = settings.GetType().GetProperty("CompatibilityOptions");
                if (compatProp != null && compatProp.CanWrite)
                {
                    var enumType = compatProp.PropertyType;
                    var enumValue = Enum.Parse(enumType, "Excel2003");
                    compatProp.SetValue(settings, enumValue);
                }

                // Refresh all pivot tables in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pivot in sheet.PivotTables)
                    {
                        try
                        {
                            // Refresh the pivot table data
                            pivot.RefreshData();

                            // If the RefreshDataOnOpen property exists, set it to true
                            var refreshProp = pivot.GetType().GetProperty("RefreshDataOnOpen");
                            if (refreshProp != null && refreshProp.CanWrite)
                            {
                                refreshProp.SetValue(pivot, true);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to refresh pivot '{pivot.Name}' on sheet '{sheet.Name}': {ex.Message}");
                        }
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
