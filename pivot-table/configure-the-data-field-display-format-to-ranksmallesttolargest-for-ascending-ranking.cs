// Title: Aspose.Cells for .NET: Set Pivot Table Data Field to Rank Smallest‑to‑Largest (Ascending)
// Description: Creates a workbook, adds sample data, builds a pivot table, and configures the data field to use PivotFieldDataDisplayFormat.RankSmallestToLargest before refreshing and saving the file.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | RankSmallestToLargest | ascending rank | pivot field display format | ShowValuesSetting | data field ranking | Excel pivot ranking example
// Common Searches: Aspose.Cells rank smallest to largest pivot | C# set PivotFieldDataDisplayFormat to RankSmallestToLargest | how to display ascending rank in Aspose.Cells pivot table | configure pivot data field ranking with Aspose.Cells .NET
// Developer Intent: Apply an ascending rank display format to a pivot table’s data field using Aspose.Cells for .NET.
// Use Cases: Generate a sales report that lists products from lowest to highest revenue. | Create a performance dashboard that highlights the smallest metrics first. | Export analytical workbooks where items are automatically ordered by minimum values for downstream processing.
// AI Prompts: Show how to set ShowValuesSetting.CalculationType to RankSmallestToLargest for a PivotField in Aspose.Cells C#. | Provide a complete .NET example that builds a pivot table, adds row and data fields, applies ascending rank display, refreshes, and saves the workbook. | Explain the effect of PivotFieldDataDisplayFormat.RankSmallestToLargest on pivot calculations and output.

using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook, adds sample data, builds a pivot table, and configures the data field to use PivotFieldDataDisplayFormat.RankSmallestToLargest before refreshing and saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].Value = "Category";
        cells["B1"].Value = "Value";
        cells["A2"].Value = "A";
        cells["B2"].Value = 10;
        cells["A3"].Value = "B";
        cells["B3"].Value = 30;
        cells["A4"].Value = "C";
        cells["B4"].Value = 20;

        // Add a pivot table covering the data range and place it at D3
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "Pivot1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add a row field (Category) and a data field (Value)
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");

        // Retrieve the data field and set its display format to rank smallest‑to‑largest
        PivotField dataField = pivot.DataFields[0];
        dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.RankSmallestToLargest;

        // Refresh the pivot table and calculate the data
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook
        workbook.Save("PivotRankSmallestToLargest.xlsx");
    }
}
