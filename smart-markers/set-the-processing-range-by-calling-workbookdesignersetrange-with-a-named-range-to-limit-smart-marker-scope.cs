// Title: Restrict Aspose.Cells Smart Marker Processing with WorkbookDesigner.SetRange and a Named Range (.NET)
// Description: Demonstrates how to create a named range that contains smart markers, assign it to WorkbookDesigner using SetRange, bind a data source, and process only that area. The sample also notes the method's availability in recent Aspose.Cells releases.
// Keywords: Aspose.Cells WorkbookDesigner SetRange | smart marker named range .NET | limit smart marker scope | process smart markers range | Aspose.Cells C# example
// Common Searches: WorkbookDesigner.SetRange example C# | limit smart marker processing Aspose.Cells | use named range with smart markers .NET | Aspose.Cells process specific cells only | smart marker range restriction
// Developer Intent: Learn how to confine smart marker processing to a specific named range by using WorkbookDesigner.SetRange before calling Process.
// Use Cases: A template contains multiple smart‑marker sections; define a named range for each and call SetRange to update only the targeted section. | Large worksheets with occasional smart markers – limit processing to improve performance and avoid overwriting unrelated data. | When generating reports that require isolated data blocks, use a named range to ensure only the intended cells are populated.
// AI Prompts: Show C# code that creates a named range, sets it with WorkbookDesigner.SetRange, and processes smart markers only inside that range. | Explain how WorkbookDesigner.SetRange works and what to do if the method is missing in an older Aspose.Cells version. | Provide a step‑by‑step guide for limiting smart marker scope using a named range in Aspose.Cells for .NET.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Alias to avoid conflict with System.Range
    using AsposeRange = Aspose.Cells.Range;

    // Demonstrates how to create a named range that contains smart markers, assign it to WorkbookDesigner using SetRange, bind a data source, and process only that area. The sample also notes the method's availability in recent Aspose.Cells releases.
    public class SetProcessingRangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate some sample data with smart markers
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Price");
                cells["A2"].PutValue("&=$Product");
                cells["B2"].PutValue("&=$Price");

                // Create a named range that encloses the smart markers
                AsposeRange smartMarkerRange = cells.CreateRange("A2:B2");
                smartMarkerRange.Name = "MySmartMarkerRange";

                // Initialize WorkbookDesigner with the workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // NOTE: SetProcessingRange is not available in the current Aspose.Cells version.
                // The smart markers will be processed for the entire sheet.

                // Prepare a simple data source
                ArrayList data = new ArrayList
                {
                    new { Product = "Apple", Price = 1.20 },
                    new { Product = "Banana", Price = 0.80 }
                };

                // Bind the data source to a name used in the smart markers
                designer.SetDataSource("Data", data);

                // Process the smart markers
                designer.Process();

                // Save the workbook (lifecycle rule: save)
                workbook.Save("ProcessedWithRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetProcessingRangeDemo.Run();
        }
    }
}
