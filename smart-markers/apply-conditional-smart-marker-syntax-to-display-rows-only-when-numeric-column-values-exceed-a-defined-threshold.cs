// Title: C# Conditional Smart Marker in Aspose.Cells – Show Rows Only When Score Exceeds a Threshold
// Description: Demonstrates how to use the &IF($Score>value) smart‑marker syntax with Aspose.Cells WorkbookDesigner. A template row is marked, a DataTable supplies employee scores, and only rows where the Score is greater than the defined limit are rendered in the final Excel file.
// Keywords: Aspose.Cells C# conditional smart marker | smart marker &IF syntax | filter rows by numeric value Aspose | WorkbookDesigner threshold example | Excel row visibility based on score | DataTable smart markers | conditional row generation Aspose.Cells
// Common Searches: Aspose.Cells hide rows with smart markers based on column value | C# &IF smart marker threshold example | How to filter Excel rows using conditional smart markers in Aspose | Conditional smart marker syntax for numeric columns | Aspose.Cells conditional row display C#
// Developer Intent: Generate an Excel workbook where rows are created only for records whose Score column exceeds a specified threshold using Aspose.Cells conditional smart markers.
// Use Cases: Performance report that lists employees with scores above a target. | Sales ledger that includes only transactions exceeding a sales quota. | Inventory list that shows items with quantity greater than the reorder level.
// AI Prompts: Write C# code with Aspose.Cells that uses &IF($Column>value) smart markers to display rows only when a numeric field meets a custom threshold. | Explain the &IF syntax in Aspose.Cells smart markers and how to make the threshold configurable at runtime. | Provide troubleshooting steps when conditional smart markers do not filter rows as expected in a WorkbookDesigner workflow.

using System;
using System.Data;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsConditionalSmartMarkerDemo
{
    // Demonstrates how to use the &IF($Score>value) smart‑marker syntax with Aspose.Cells WorkbookDesigner. A template row is marked, a DataTable supplies employee scores, and only rows where the Score is greater than the defined limit are rendered in the final Excel file.
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

                // Add header row
                cells["A1"].PutValue("Name");
                cells["B1"].PutValue("Score");

                // Row 2 acts as the template row for conditional smart markers.
                // The syntax &IF($Score>80) ... &ENDIF ensures the row is displayed only when Score > 80.
                cells["A2"].PutValue("&IF($Score>80)&=$Name&ENDIF");
                cells["B2"].PutValue("&IF($Score>80)&=$Score&ENDIF");

                // Define the range that contains the smart markers.
                // Naming the range "_CellsSmartMarkers" tells the designer to process it.
                AsposeRange smartRange = cells.CreateRange("A2:B2");
                smartRange.Name = "_CellsSmartMarkers";

                // Prepare sample data in a DataTable
                DataTable dt = new DataTable("Employees");
                dt.Columns.Add("Name", typeof(string));
                dt.Columns.Add("Score", typeof(int));

                dt.Rows.Add("Alice", 95);   // Should appear
                dt.Rows.Add("Bob", 67);     // Should be hidden
                dt.Rows.Add("Charlie", 82); // Should appear
                dt.Rows.Add("Diana", 74);   // Should be hidden

                // Set the data source for the designer
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource(dt);

                // Process the smart markers
                designer.Process();

                // Save the resulting workbook
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
