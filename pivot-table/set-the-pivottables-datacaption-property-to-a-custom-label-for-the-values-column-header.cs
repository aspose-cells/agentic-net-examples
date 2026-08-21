// Title: Aspose.Cells C# – Set a custom Values column header in a PivotTable (DataFieldHeaderName)
// Description: Shows how to create a workbook, fill it with sample sales data, add a PivotTable, assign 'Product' to rows and 'Sales' to the data area, then rename the default Values column by setting PivotTable.DataFieldHeaderName to a custom string before refreshing and saving the file.
// Keywords: Aspose.Cells | C# | PivotTable | DataFieldHeaderName | custom values header | rename pivot table column | Excel export | Aspose.Cells for .NET example | pivot table header customization | GitHub code sample
// Common Searches: Aspose.Cells set custom header for pivot table values column | How to rename Values column in Aspose.Cells PivotTable C# | PivotTable DataFieldHeaderName property example | Change pivot table data caption Aspose.Cells .NET | C# code to customize pivot table header Aspose
// Developer Intent: Rename the PivotTable values column header to a user‑defined label.
// Use Cases: Generate a sales report where the data column reads "Total Sales" instead of the generic "Values". | Create a regional performance dashboard with the data field header changed to "Revenue" to align with corporate terminology. | Automate Excel exports that include a pivot table with a dynamic caption such as "Custom Values" defined at runtime.
// AI Prompts: Write C# code using Aspose.Cells that builds a PivotTable from a range and sets a custom string for the values column header. | Show how to apply the PivotTable.DataFieldHeaderName property to rename the data field caption in an Aspose.Cells workbook. | Provide a step‑by‑step example: create workbook, add sample data, create PivotTable, add row and data fields, customize the values column header, refresh, and save.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, fill it with sample sales data, add a PivotTable, assign 'Product' to rows and 'Sales' to the data area, then rename the default Values column by setting PivotTable.DataFieldHeaderName to a custom string before refreshing and saving the file.
    public class SetPivotDataFieldHeaderDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Product";
                sheet.Cells["B1"].Value = "Region";
                sheet.Cells["C1"].Value = "Sales";

                sheet.Cells["A2"].Value = "Laptop";
                sheet.Cells["B2"].Value = "North";
                sheet.Cells["C2"].Value = 1000;

                sheet.Cells["A3"].Value = "Laptop";
                sheet.Cells["B3"].Value = "South";
                sheet.Cells["C3"].Value = 1500;

                sheet.Cells["A4"].Value = "Phone";
                sheet.Cells["B4"].Value = "North";
                sheet.Cells["C4"].Value = 800;

                sheet.Cells["A5"].Value = "Phone";
                sheet.Cells["B5"].Value = "South";
                sheet.Cells["C5"].Value = 1200;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: rows and data fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Set a custom caption for the values column header (data field header)
                pivotTable.DataFieldHeaderName = "Custom Values";

                // Refresh the pivot cache and calculate the pivot table
                pivotTable.RefreshData();      // Correct method to refresh cache
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTableCustomDataHeader.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
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
            SetPivotDataFieldHeaderDemo.Run();
        }
    }
}
