// Title: Enable BackgroundRefresh for WebQuery tables in an Excel workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells set BackgroundRefresh true for WebQueryTables to prevent UI freeze | C# load workbook and enable asynchronous web query refresh using Aspose.Cells | How to use reflection to access WebQueryTables BackgroundRefresh property in Aspose.Cells | Avoid blocking UI when refreshing web data connections in Excel via Aspose.Cells | Enable background data retrieval for web queries in .NET Excel automation
// Tags: Aspose.Cells background refresh web query | C# enable asynchronous web query Excel | reflection access WebQueryTables property Aspose.Cells | prevent UI blocking Excel data connection .NET | set BackgroundRefresh property Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections;
using System.IO;

// The example loads an existing Excel workbook with Aspose.Cells, iterates through each worksheet, uses reflection to locate the WebQueryTables collection, and sets each table's BackgroundRefresh property to true. This enables asynchronous data retrieval for web queries, preventing UI blocking, and then saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate over worksheets and enable background refresh for any web query tables (using reflection for compatibility)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var webQueryTablesProp = sheet.GetType().GetProperty("WebQueryTables");
                if (webQueryTablesProp == null) continue; // Property not available in this version

                var webQueryTables = webQueryTablesProp.GetValue(sheet) as IEnumerable;
                if (webQueryTables == null) continue;

                foreach (var webQueryObj in webQueryTables)
                {
                    var bgRefreshProp = webQueryObj.GetType().GetProperty("BackgroundRefresh");
                    if (bgRefreshProp != null && bgRefreshProp.CanWrite)
                    {
                        bgRefreshProp.SetValue(webQueryObj, true);
                    }
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Output exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
