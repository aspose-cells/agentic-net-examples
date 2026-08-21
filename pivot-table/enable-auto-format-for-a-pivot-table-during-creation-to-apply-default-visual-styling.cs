// Title: Apply Default Visual Style to a New PivotTable with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, add sample data, insert a PivotTable, enable the IsAutoFormat property to apply the built‑in style automatically, calculate the data, and save the file.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | IsAutoFormat | automatic formatting | default style | Excel report | programmatic pivot styling
// Common Searches: Aspose.Cells enable pivot table auto format | C# set IsAutoFormat true | default pivot table style Aspose.Cells | apply built‑in style to PivotTable programmatically | Aspose.Cells pivot table formatting options
// Developer Intent: Create a PivotTable that automatically receives Excel’s built‑in visual style without manual styling code.
// Use Cases: Generate analytical reports where every PivotTable uses a consistent default style. | Loop through multiple data sets, create PivotTables, and enable IsAutoFormat to ensure uniform appearance. | Export business intelligence dashboards to Excel with ready‑to‑read, auto‑styled PivotTables.
// AI Prompts: How do I enable automatic formatting for a PivotTable using Aspose.Cells in C#? | Provide C# code that creates several PivotTables, sets IsAutoFormat = true for each, calculates the data, and saves the workbook. | Explain the impact of the IsAutoFormat property on PivotTable appearance and how to change the default style if needed.

using Aspose.Cells;
using Aspose.Cells.Pivot;

// Shows how to create a workbook, add sample data, insert a PivotTable, enable the IsAutoFormat property to apply the built‑in style automatically, calculate the data, and save the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(200);
        worksheet.Cells["A4"].PutValue("A");
        worksheet.Cells["B4"].PutValue(150);

        // Add a pivot table to the worksheet
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

        // Enable automatic formatting (default visual styling)
        pivotTable.IsAutoFormat = true;

        // Calculate the pivot table data
        pivotTable.CalculateData();

        // Save the workbook with the formatted pivot table
        workbook.Save("PivotTableAutoFormatDemo.xlsx");
    }
}
