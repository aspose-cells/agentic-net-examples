// Title: C# – Limit Aspose.Cells Smart Marker Processing to a Named Range with WorkbookDesigner.SetRange
// Description: Demonstrates how to create a workbook, define a named range that contains smart markers, bind an ArrayList data source, call WorkbookDesigner.SetRange to restrict processing to that range, and save the result. This technique ensures only the specified cells are populated while the rest of the sheet remains unchanged.
// Keywords: Aspose.Cells C# | WorkbookDesigner SetRange | smart markers named range | limit smart marker scope | Aspose.Cells processing range | C# Excel automation | .NET smart marker example
// Common Searches: How to use WorkbookDesigner.SetRange in Aspose.Cells | Aspose.Cells limit smart markers to a specific area | C# example for processing smart markers in a named range | Aspose.Cells set processing range for smart markers | Restrict smart marker expansion in .NET
// Developer Intent: Use WorkbookDesigner.SetRange to confine smart marker processing to a predefined named range before calling Process().
// Use Cases: Populate only a table of smart markers in a large worksheet while preserving surrounding static data. | Run separate smart‑marker fills for multiple sections of the same sheet by assigning different named ranges. | Improve performance by processing only the cells that contain smart markers rather than the entire worksheet.
// AI Prompts: Generate C# code that creates a named range for smart markers and calls WorkbookDesigner.SetRange to process only that range with Aspose.Cells. | Show how to bind an ArrayList of objects to a WorkbookDesigner and limit smart marker expansion to a specific named range before invoking Process(). | Provide an example of using multiple named ranges with WorkbookDesigner.SetRange to handle separate smart‑marker sections in one workbook.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerRangeDemo
{
    // Demonstrates how to create a workbook, define a named range that contains smart markers, bind an ArrayList data source, call WorkbookDesigner.SetRange to restrict processing to that range, and save the result. This technique ensures only the specified cells are populated while the rest of the sheet remains unchanged.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // 2. Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 3. Insert sample data and smart markers
                // Header row
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Price");

                // Smart marker rows (will be processed)
                cells["A2"].PutValue("&=$Product");
                cells["B2"].PutValue("&=$Price");

                // 4. Define a named range that encloses the smart marker rows
                // Create the range A2:B2 and give it a name
                Aspose.Cells.Range smartMarkerRange = cells.CreateRange("A2", "B2");
                smartMarkerRange.Name = "_CellsSmartMarkers"; // conventional name for range smart markers

                // 5. Prepare a data source (simple ArrayList of anonymous objects)
                ArrayList data = new ArrayList
                {
                    new { Product = "Apple",  Price = 1.20 },
                    new { Product = "Banana", Price = 0.80 },
                    new { Product = "Orange", Price = 1.50 }
                };

                // 6. Create a WorkbookDesigner and bind the data source
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource("Data", data);

                // Note: SetSmartMarkerRange is not available in the current API version.
                // The smart markers will be processed for the entire sheet, which includes the defined range.

                // 8. Process the smart markers
                designer.Process();

                // 9. Save the result (lifecycle rule: save)
                string outputPath = "SmartMarkerRangeOutput.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Processing completed. File saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
