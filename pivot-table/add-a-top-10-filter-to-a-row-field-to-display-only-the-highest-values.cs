// Title: Apply a Top N filter to a PivotTable row field with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, inserts sample sales data, builds a PivotTable (Category as rows, Sales as values), and uses PivotTable.BaseFields[0].FilterTop10 to display only the highest‑ranking categories (e.g., top 3 by sum). The pivot is refreshed and saved as an .xlsx file.
// Keywords: Aspose.Cells | C# | PivotTable | FilterTop10 | Top N filter | row field filter | sum aggregation | Excel automation | pivot top values | Aspose.Cells .NET
// Common Searches: Aspose.Cells top N filter pivot table | C# FilterTop10 method example | show top categories in Aspose.Cells pivot | apply top 10 filter to pivot rows using Aspose.Cells | PivotTable row filter highest values .NET
// Developer Intent: The developer wants to limit a PivotTable row field to the highest‑valued items (Top N) using Aspose.Cells in C#.
// Use Cases: Generate a sales report that automatically lists only the top‑selling product categories. | Create a dashboard that highlights the top N regions or stores based on revenue. | Produce a concise summary sheet that filters out low‑performing items from a pivot analysis.
// AI Prompts: Write C# code with Aspose.Cells to add a PivotTable and apply a Top N filter on the row field using the sum of a data field. | Show how to change the item count in FilterTop10 to display the top 5 rows instead of 3. | Explain how to combine FilterTop10 with a label filter in Aspose.Cells for .NET. | Provide a step‑by‑step guide to refresh the pivot after applying a Top N filter.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook, inserts sample sales data, builds a PivotTable (Category as rows, Sales as values), and uses PivotTable.BaseFields[0].FilterTop10 to display only the highest‑ranking categories (e.g., top 3 by sum). The pivot is refreshed and saved as an .xlsx file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].Value = "Category";
        worksheet.Cells["B1"].Value = "Sales";

        worksheet.Cells["A2"].Value = "Fruit";
        worksheet.Cells["B2"].Value = 120;

        worksheet.Cells["A3"].Value = "Vegetable";
        worksheet.Cells["B3"].Value = 80;

        worksheet.Cells["A4"].Value = "Dairy";
        worksheet.Cells["B4"].Value = 150;

        worksheet.Cells["A5"].Value = "Meat";
        worksheet.Cells["B5"].Value = 200;

        worksheet.Cells["A6"].Value = "Bakery";
        worksheet.Cells["B6"].Value = 60;

        // Create a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B6", "D3", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Add the Category field as a row field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

        // Add the Sales field as a data field
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Apply a Top 10 filter on the row field to show only the highest values.
        // Here we show the top 3 items based on the Sum of Sales.
        // Parameters: valueFieldIndex = 0 (first data field), type = Sum, isTop = true, itemCount = 3
        pivotTable.BaseFields[0].FilterTop10(0, PivotFilterType.Sum, true, 3);

        // Refresh the pivot table data
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("PivotTop10Filter.xlsx");
    }
}
