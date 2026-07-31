// Title: Set PivotTable Data Field to Sum with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills a simple data set, adds a PivotTable at D3, assigns "Category" as a row field, adds "Amount" as a data field, configures its ConsolidationFunction to Sum, refreshes and calculates the pivot, then saves the file as PivotTableWithSumFunction.xlsx.
// Keywords: Aspose.Cells PivotTable Sum C# | ConsolidationFunction.Sum example | Aspose.Cells set data field aggregation | C# pivot table sum function | Aspose.Cells .NET pivot table tutorial
// Common Searches: Aspose.Cells set pivot table consolidation function to sum | C# example for summing values in Aspose.Cells pivot table | How to use ConsolidationFunction.Sum in Aspose.Cells | PivotTable data field aggregation Aspose.Cells .NET
// Developer Intent: Configure a PivotTable data field to aggregate values using the Sum function in Aspose.Cells for .NET.
// Use Cases: Generate sales reports that total amounts per category automatically. | Create financial dashboards that sum transaction values without manual formulas. | Build dynamic Excel exports where pivot tables summarize large data sets with a single Sum aggregation.
// AI Prompts: Write C# code with Aspose.Cells to add a PivotTable and set its data field's ConsolidationFunction to Sum. | Explain the effect of ConsolidationFunction.Sum on PivotTable calculations and the required refresh steps in Aspose.Cells. | Provide a step‑by‑step guide for creating a PivotTable in Aspose.Cells and configuring a data field for summation.

using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook, fills a simple data set, adds a PivotTable at D3, assigns "Category" as a row field, adds "Amount" as a data field, configures its ConsolidationFunction to Sum, refreshes and calculates the pivot, then saves the file as PivotTableWithSumFunction.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].Value = "Category";
        cells["B1"].Value = "Amount";
        cells["A2"].Value = "A";
        cells["B2"].Value = 100;
        cells["A3"].Value = "B";
        cells["B3"].Value = 150;
        cells["A4"].Value = "A";
        cells["B4"].Value = 200;
        cells["A5"].Value = "B";
        cells["B5"].Value = 120;

        // Add a pivot table using the data range A1:B5, placing it at D3
        PivotTableCollection pivotTables = worksheet.PivotTables;
        int pivotIndex = pivotTables.Add("A1:B5", "D3", "PivotTableSum");
        PivotTable pivotTable = pivotTables[pivotIndex];

        // Add the Category field as a row field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

        // Add the Amount field as a data field and set its consolidation function to Sum
        int dataFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
        PivotField amountField = pivotTable.DataFields[dataFieldIndex];
        amountField.Function = ConsolidationFunction.Sum; // Set aggregation to Sum

        // Refresh the pivot table data and calculate the results
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the configured pivot table
        workbook.Save("PivotTableWithSumFunction.xlsx");
    }
}
