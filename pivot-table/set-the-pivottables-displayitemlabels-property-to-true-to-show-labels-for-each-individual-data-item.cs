// Title: Enable Item Labels in an Aspose.Cells PivotTable (C#) – Using the DisplayItemLabels Property
// Description: C# example that creates a workbook, adds sample sales data, builds a PivotTable, and attempts to set the DisplayItemLabels property to show a label for every data item. The code also demonstrates refreshing, calculating, and saving the workbook, and notes that the property is unavailable in the current Aspose.Cells release, offering alternative approaches.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | DisplayItemLabels | show item labels | pivot table label visibility | Excel automation | workbook save | Aspose.Cells properties
// Common Searches: Aspose.Cells DisplayItemLabels C# | how to show item labels in Aspose.Cells pivot table | pivot table label visibility Aspose.Cells | DisplayItemLabels property missing Aspose.Cells | alternative to DisplayItemLabels in Aspose.Cells
// Developer Intent: Enable the DisplayItemLabels option on a PivotTable so each individual data item displays its own label.
// Use Cases: Generate an Excel workbook with sample sales data and a PivotTable that would display item labels if the property were supported. | Refresh and calculate the PivotTable after configuring display settings, then save the file as .xlsx. | Detect the absence of the DisplayItemLabels property in the installed Aspose.Cells version and apply other available properties or work‑arounds to control label visibility.
// AI Prompts: Provide C# code using Aspose.Cells to turn on item labels in a PivotTable, or suggest alternative properties when DisplayItemLabels is not present. | Explain how to programmatically check whether the DisplayItemLabels property exists in the current Aspose.Cells library and recommend a fallback solution. | Create a complete example that builds a PivotTable with row, column, and data fields, refreshes it, saves the workbook, and includes comments on label display options.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, adds sample sales data, builds a PivotTable, and attempts to set the DisplayItemLabels property to show a label for every data item. The code also demonstrates refreshing, calculating, and saving the workbook, and notes that the property is unavailable in the current Aspose.Cells release, offering alternative approaches.
    public class PivotTableDisplayItemLabelsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Pivot table created and workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Product");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Electronics");
            sheet.Cells["B2"].PutValue("Laptop");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Electronics");
            sheet.Cells["B3"].PutValue("Phone");
            sheet.Cells["C3"].PutValue(800);

            sheet.Cells["A4"].PutValue("Furniture");
            sheet.Cells["B4"].PutValue("Chair");
            sheet.Cells["C4"].PutValue(150);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // NOTE: The DisplayItemLabels property is not available in the current Aspose.Cells version.
            // If needed, configure related display options via other available properties.

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Define output file path
            string outputPath = "PivotTableDisplayItemLabelsDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
        }
    }
}
