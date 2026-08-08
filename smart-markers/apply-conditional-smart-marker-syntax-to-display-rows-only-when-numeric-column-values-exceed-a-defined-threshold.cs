// Title: C# Aspose.Cells Conditional Smart Marker – Display Rows When Amount Exceeds Threshold
// Description: Shows how to use Aspose.Cells WorkbookDesigner with &IF…&ENDIF smart‑marker syntax and a named _CellsSmartMarkers range to include only rows where the numeric Amount column is greater than a defined limit.
// Keywords: Aspose.Cells | C# smart markers | conditional smart marker | Excel row filter | &IF syntax | WorkbookDesigner | range smart markers | .NET Excel export | numeric comparison | threshold filter
// Common Searches: Aspose.Cells conditional smart marker C# | filter rows with smart markers Excel | use &IF in Aspose.Cells template | named range _CellsSmartMarkers example | show rows only if value > 100 Aspose | smart marker numeric comparison .NET
// Developer Intent: Generate an Excel workbook that includes only rows where a numeric column satisfies a specified condition using Aspose.Cells smart markers.
// Use Cases: Sales report that lists products with revenue above a set amount | Inventory sheet that omits items below the reorder quantity | Financial ledger displaying transactions exceeding a monetary threshold | Customer list filtered by minimum purchase value
// AI Prompts: Generate code that reads the threshold from a configuration file and applies it in the &IF condition. | Show how to combine multiple &IF conditions (e.g., Amount > 100 AND Category = "Fruit") in a single smart‑marker block. | Provide an example that writes the processed workbook to a MemoryStream and returns it as a byte array.

using System;
using System.Data;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace ConditionalSmartMarkerDemo
{
    // Shows how to use Aspose.Cells WorkbookDesigner with &IF…&ENDIF smart‑marker syntax and a named _CellsSmartMarkers range to include only rows where the numeric Amount column is greater than a defined limit.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ----- Template setup -----
                // Header row
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Amount");

                // Smart marker row with conditional display:
                // &IF($Amount>100)   -> start condition (show row only if Amount > 100)
                // &=$Product         -> populate Product column
                // &=$Amount          -> populate Amount column
                // &ENDIF             -> end condition
                cells["A2"].PutValue("&IF($Amount>100)");
                cells["A3"].PutValue("&=$Product");
                cells["B3"].PutValue("&=$Amount");
                cells["A4"].PutValue("&ENDIF");

                // Define the range that contains the smart markers
                // The range must be named "_CellsSmartMarkers" when using range smart markers
                AsposeRange smRange = sheet.Cells.CreateRange("A2:B4");
                smRange.Name = "_CellsSmartMarkers";

                // ----- Data source creation -----
                DataTable dt = new DataTable("Products");
                dt.Columns.Add("Product", typeof(string));
                dt.Columns.Add("Amount", typeof(double));

                // Sample data: only rows with Amount > 100 will be displayed
                dt.Rows.Add("Apple", 80);    // will be hidden
                dt.Rows.Add("Banana", 150);  // will be shown
                dt.Rows.Add("Cherry", 200);  // will be shown
                dt.Rows.Add("Date", 50);     // will be hidden

                // ----- Designer configuration -----
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Set the data source (the table name must match the smart marker table name)
                designer.SetDataSource("Products", dt);

                // Process the smart markers
                designer.Process();

                // ----- Save the result -----
                string outputPath = "ConditionalSmartMarkerResult.xlsx";
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
