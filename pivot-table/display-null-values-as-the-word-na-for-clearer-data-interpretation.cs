// Title: Aspose.Cells C# Pivot Table – Show Null Cells as “N/A”
// Description: C# code that builds a workbook, inserts product‑sales data with null entries, creates a pivot table, and configures it to display “N/A” for any null values by enabling DisplayNullString and setting NullString. The pivot cache is refreshed, data calculated, and the workbook saved as PivotTable_NullAsNA.xlsx.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | DisplayNullString | NullString | N/A | null handling | Excel export | data analysis
// Common Searches: Aspose.Cells show N/A for null values | C# pivot table null string Aspose | DisplayNullString property example | Set NullString in Aspose.Cells pivot | Replace empty cells with N/A in Excel using Aspose
// Developer Intent: Replace empty or null cells in an Aspose.Cells pivot table with the text “N/A” using .NET.
// Use Cases: Generate an Excel workbook with sample data that includes missing product names or sales figures. | Create a pivot table that summarizes sales by product while presenting missing values as “N/A”. | Apply DisplayNullString = true and NullString = "N/A" to customize null representation, then refresh and calculate the pivot cache before saving.
// AI Prompts: Write C# code using Aspose.Cells to create a pivot table that displays "N/A" for null entries. | Explain how DisplayNullString and NullString properties work in Aspose.Cells PivotTable objects. | Provide step‑by‑step instructions to refresh and calculate a pivot table after setting a custom null display string in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // C# code that builds a workbook, inserts product‑sales data with null entries, creates a pivot table, and configures it to display “N/A” for any null values by enabling DisplayNullString and setting NullString. The pivot cache is refreshed, data calculated, and the workbook saved as PivotTable_NullAsNA.xlsx.
    public class DisplayNullAsNA
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with null values
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(1500);
                sheet.Cells["A4"].PutValue(null);      // Null product name
                sheet.Cells["B4"].PutValue(null);      // Null sales value
                sheet.Cells["A5"].PutValue("Banana");
                sheet.Cells["B5"].PutValue(800);

                // Add a pivot table covering the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: rows = Product, data = Sales
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

                // Enable custom display for null values and set the desired string
                pivotTable.DisplayNullString = true;
                pivotTable.NullString = "N/A";

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();   // Correct method to refresh pivot cache
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTable_NullAsNA.xlsx");
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
            DisplayNullAsNA.Run();
        }
    }
}
