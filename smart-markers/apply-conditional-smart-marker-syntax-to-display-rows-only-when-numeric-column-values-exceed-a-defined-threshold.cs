// Title: Use Aspose.Cells C# conditional smart markers to display rows only when a numeric column exceeds a specified threshold
// AI Prompts: Create a C# program that builds an Aspose.Cells workbook, defines a smart‑marker block, and uses the &IF($Value>limit) construct to add rows only when the numeric field exceeds the limit. | Demonstrate how to bind a List<Item> to WorkbookDesigner and apply a conditional smart marker that suppresses rows where the Value property is less than or equal to a given threshold.
// Common Searches: aspnet aspocells how to hide rows with &IF smart marker based on column value | c# example of conditional smart marker that only prints rows where numeric column > 50 | using Aspose.Cells WorkbookDesigner to generate rows conditionally from a collection | smart marker syntax &IF($Amount>threshold) for Excel export in C# | filtering Excel rows with Aspose.Cells smart markers and numeric thresholds
// Tags: Aspose.Cells &IF smart marker | WorkbookDesigner data binding | smart marker conditional expression | Excel row filtering by numeric value | C# range smart marker definition

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsConditionalSmartMarkerDemo
{
    // Simple data class
    // The example creates a new workbook, defines a smart‑marker range, inserts an &IF($Value>50) condition to include rows only when the Value exceeds 50, binds a List<Item> to WorkbookDesigner, processes the markers so qualifying rows are generated, and saves the result as an Excel file.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public double Value { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set up column headers
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["B1"].PutValue("Value");

                // Define the smart marker range (the row that will be repeated)
                // The range name "_CellsSmartMarkers" tells the designer to treat it as a smart marker block
                Aspose.Cells.Range smartRange = sheet.Cells.CreateRange("A2:B2");
                smartRange.Name = "_CellsSmartMarkers";

                // Insert conditional smart markers.
                // The row will be populated only when the numeric column "Value" exceeds the threshold (e.g., 50).
                // &IF($Value>50) starts the condition, &ENDIF ends it.
                sheet.Cells["A2"].PutValue("&IF($Value>50)&=$Name");
                sheet.Cells["B2"].PutValue("&=$Value&ENDIF");

                // Prepare sample data
                List<Item> items = new List<Item>
                {
                    new Item { Name = "Alpha",   Value = 30 }, // Will be hidden (30 <= 50)
                    new Item { Name = "Beta",    Value = 75 }, // Will be shown
                    new Item { Name = "Gamma",   Value = 55 }, // Will be shown
                    new Item { Name = "Delta",   Value = 20 }  // Will be hidden
                };

                // Set up the workbook designer and bind the data source
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                    // LineByLine is obsolete; not required when using range smart markers
                };
                designer.SetDataSource("Items", items);

                // Process the smart markers (the range will be expanded only for rows meeting the condition)
                designer.Process();

                // Save the result
                string outputPath = "ConditionalSmartMarkerOutput.xlsx";
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
