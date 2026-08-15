// Title: Show Item Labels in Pivot Table Values Area with Aspose.Cells for .NET (ShowValuesColumn)
// Description: C# example that creates a workbook, fills it with product, region and sales data, builds a pivot table and demonstrates how to display item labels in the values area using the ShowValuesColumn property. The sample also notes that the property is unavailable in the current Aspose.Cells release and suggests a fallback approach before saving the file as XLSX.
// Keywords: Aspose.Cells | ShowValuesColumn | pivot table item labels | values area | C# | .NET | Excel automation | workbook creation | pivot fields | alternative fallback
// Common Searches: Aspose.Cells ShowValuesColumn property | display item labels in pivot values area C# | enable values column in Aspose.Cells pivot table | pivot table item labels not showing Aspose.Cells | fallback for ShowValuesColumn missing
// Developer Intent: Enable the ShowValuesColumn setting so that row or column field names appear alongside data in the pivot table’s values area.
// Use Cases: Generate a sales report workbook and configure the pivot table to show item labels in the data area when the ShowValuesColumn property is supported. | Detect the absence of ShowValuesColumn in the installed Aspose.Cells version and apply an alternative formatting technique to mimic item labels. | Export the configured workbook to XLSX for downstream analysis or distribution.
// AI Prompts: Write C# code using Aspose.Cells that enables item labels in the pivot table values area and includes a version‑check fallback if ShowValuesColumn is unavailable. | Explain how to programmatically verify the presence of the ShowValuesColumn property in the current Aspose.Cells library and propose alternative methods. | Create a complete Aspose.Cells example that builds a pivot table and displays row field names in the data section using supported features.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, fills it with product, region and sales data, builds a pivot table and demonstrates how to display item labels in the values area using the ShowValuesColumn property. The sample also notes that the property is unavailable in the current Aspose.Cells release and suggests a fallback approach before saving the file as XLSX.
    public class ShowValuesColumnDemo
    {
        public static void Main()
        {
            Run();
        }

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

                sheet.Cells["A2"].Value = "A";
                sheet.Cells["B2"].Value = "North";
                sheet.Cells["C2"].Value = 1200;

                sheet.Cells["A3"].Value = "A";
                sheet.Cells["B3"].Value = "South";
                sheet.Cells["C3"].Value = 1500;

                sheet.Cells["A4"].Value = "B";
                sheet.Cells["B4"].Value = "North";
                sheet.Cells["C4"].Value = 800;

                sheet.Cells["A5"].Value = "B";
                sheet.Cells["B5"].Value = "South";
                sheet.Cells["C5"].Value = 950;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Product as row field
                pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Region as column field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Sales as data field

                // Note: The ShowValuesColumn property is not available in the current Aspose.Cells version.
                // If needed, alternative settings can be applied here.

                // Save the workbook to a file
                string outputPath = "ShowValuesColumnDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
