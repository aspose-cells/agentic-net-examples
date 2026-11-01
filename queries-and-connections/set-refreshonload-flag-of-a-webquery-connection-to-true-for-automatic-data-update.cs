// Title: Set RefreshOnLoad = true for every WebQuery table in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set RefreshOnLoad property for WebQuery tables | How to enable automatic refresh of web query connections in an Excel file with Aspose.Cells | Iterate worksheets and update WebQueryTables RefreshOnLoad flag using reflection in C# | Programmatically turn on RefreshOnLoad for Excel web queries with Aspose.Cells .NET | C# code to enable RefreshOnLoad for all WebQuery connections in a workbook
// Tags: Aspose.Cells RefreshOnLoad WebQueryTables | C# update Excel web query connection property | programmatic enable web query auto refresh Aspose | iterate worksheets set WebQuery RefreshOnLoad | reflection based access to WebQueryTables in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, loops through each worksheet, uses reflection to obtain the WebQueryTables collection, sets each table's RefreshOnLoad property to true, and saves the workbook, handling missing files and runtime errors.
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
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and enable RefreshOnLoad for any WebQuery tables
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Use reflection to access WebQueryTables collection (avoids compile‑time dependency)
                    var webQueryTablesProp = sheet.GetType().GetProperty("WebQueryTables");
                    var webQueryTables = webQueryTablesProp?.GetValue(sheet) as System.Collections.IEnumerable;

                    if (webQueryTables != null)
                    {
                        foreach (var wqt in webQueryTables)
                        {
                            var refreshProp = wqt.GetType().GetProperty("RefreshOnLoad");
                            if (refreshProp != null && refreshProp.CanWrite)
                            {
                                refreshProp.SetValue(wqt, true);
                            }
                        }
                    }
                }
                catch (Exception innerEx)
                {
                    // Log but continue processing other worksheets
                    Console.WriteLine($"Warning: Could not process WebQuery tables on sheet '{sheet.Name}': {innerEx.Message}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
