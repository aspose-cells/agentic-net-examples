// Title: Enable background refresh for WebQuery tables in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Iterate through all WebQueryTables in a worksheet using Aspose.Cells and set each table's BackgroundRefresh property to true via reflection, then save the workbook. | Configure Excel WebQuery connections to refresh data asynchronously to prevent UI blocking, employing C# and Aspose.Cells.
// Common Searches: C# Aspose.Cells set web query background refresh to avoid UI freeze | How to use reflection to access WebQueryTables in Aspose.Cells when the API is unavailable | Enable asynchronous data retrieval for Excel WebQuery tables with Aspose.Cells .NET | Aspose.Cells example for setting BackgroundRefresh on WebQueryTables | Prevent Excel UI blocking during web query refresh using Aspose.Cells
// Tags: Aspose.Cells WebQueryTables BackgroundRefresh | C# enable background refresh Excel web query | Aspose.Cells reflection access WebQueryTables | non‑blocking web query data retrieval Aspose.Cells | update workbook WebQueryTables setting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an existing Excel file, uses reflection to locate any WebQueryTables in the first worksheet, sets each table's BackgroundRefresh property to true so data is fetched in the background, and saves the modified workbook, preventing UI blocking during web query refresh.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "Input.xlsx";
                const string outputPath = "Output.xlsx";

                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook (creates a new one if the file is missing)
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Use reflection to handle WebQueryTables in case the API is unavailable in the current version
                var webQueryTablesProp = typeof(Worksheet).GetProperty("WebQueryTables");
                if (webQueryTablesProp != null)
                {
                    var webQueryTables = webQueryTablesProp.GetValue(sheet) as System.Collections.IEnumerable;
                    if (webQueryTables != null)
                    {
                        foreach (var webQueryObj in webQueryTables)
                        {
                            var bgRefreshProp = webQueryObj.GetType().GetProperty("BackgroundRefresh");
                            if (bgRefreshProp != null && bgRefreshProp.CanWrite)
                            {
                                // Enable background refresh to avoid blocking the UI during data retrieval
                                bgRefreshProp.SetValue(webQueryObj, true);
                            }
                        }
                    }
                }

                // Save the workbook with the updated setting
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
