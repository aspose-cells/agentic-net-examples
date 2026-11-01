// Title: Disable multiple filter selections in an Aspose.Cells PivotTable using C# (AllowMultipleFilters = false)
// Common Searches: C# code to prevent users from selecting multiple items in an Aspose.Cells PivotTable filter | How to set AllowMultipleFilters to false for all fields in an Aspose.Cells PivotTable | Aspose.Cells .NET disable report filter multi‑select in a PivotTable | Restrict PivotTable row and column filters to single selection with Aspose.Cells | Using reflection to set the AllowMultipleFilters property in older Aspose.Cells versions
// Tags: Aspose.Cells pivot filter single-select | C# modify PivotField AllowMultipleFilters | Aspose.Cells reflection property assignment | PivotTable row and column filter restriction | Aspose.Cells workbook save after pivot update

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable and related classes

namespace PivotTableExample
{
    // The example loads a workbook, accesses its first PivotTable, and uses reflection to set the AllowMultipleFilters property to false on all row, column, and page fields, thereby limiting filter selections to a single item before saving the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the existing workbook that contains a PivotTable
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the first PivotTable in the worksheet, if any
                if (worksheet.PivotTables.Count > 0)
                {
                    PivotTable pivotTable = worksheet.PivotTables[0];

                    // Helper to set AllowMultipleFilters via reflection (property may not exist in older versions)
                    void TryDisableMultipleFilters(object fieldObj)
                    {
                        try
                        {
                            var prop = fieldObj.GetType().GetProperty("AllowMultipleFilters");
                            if (prop != null && prop.CanWrite)
                            {
                                prop.SetValue(fieldObj, false);
                            }
                        }
                        catch
                        {
                            // Silently ignore if the property cannot be set
                        }
                    }

                    // Restrict filter selections for end users by disabling multiple filters
                    // on all row fields
                    foreach (PivotField field in pivotTable.RowFields)
                    {
                        TryDisableMultipleFilters(field);
                    }

                    // ...and on all column fields
                    foreach (PivotField field in pivotTable.ColumnFields)
                    {
                        TryDisableMultipleFilters(field);
                    }

                    // If there are page fields (report filters), also restrict them
                    foreach (var pageFieldObj in pivotTable.PageFields)
                    {
                        // In some versions PageFields returns PivotField directly,
                        // in others it may return a different type; handle both.
                        TryDisableMultipleFilters(pageFieldObj);
                    }

                    Console.WriteLine("PivotTable filter restrictions applied (where supported).");
                }
                else
                {
                    Console.WriteLine("No PivotTable found in the worksheet.");
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
