// Title: Copy a Worksheet with Its Pivot Table Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add sample data and a pivot table, duplicate the worksheet with AddCopy, rename the copy, refresh its pivot tables, and save the file so the cloned sheet retains a functional pivot for further analysis.
// Keywords: Aspose.Cells copy worksheet | duplicate sheet retain pivot | AddCopy pivot table .NET | RefreshPivotTables C# | clone worksheet with pivot | Aspose.Cells worksheet duplication
// Common Searches: copy worksheet with pivot table Aspose.Cells | how to retain pivot tables when duplicating a sheet in C# | AddCopy method pivot table refresh | duplicate Excel sheet programmatically Aspose.Cells | clone sheet and keep pivot functionality
// Developer Intent: Duplicate an existing worksheet while preserving and refreshing its pivot tables for immediate use in analysis or reporting.
// Use Cases: Create a master sales sheet with a pivot, copy it for each regional team, and refresh the pivots to generate localized reports. | Clone a monthly budgeting worksheet that contains a pivot, modify the data in the copy, and update the pivot for scenario planning. | Generate template-driven dashboards where each duplicated sheet keeps the original pivot layout, allowing independent data manipulation.
// AI Prompts: Generate C# code that uses Aspose.Cells to copy a worksheet and automatically refresh any pivot tables in the duplicated sheet. | Explain the AddCopy method in Aspose.Cells, show how to rename the copied worksheet, refresh its pivots, and save the workbook. | Provide a step‑by‑step tutorial for cloning a worksheet containing a pivot table, updating the pivot, and exporting the result to an Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook, add sample data and a pivot table, duplicate the worksheet with AddCopy, rename the copy, refresh its pivot tables, and save the file so the cloned sheet retains a functional pivot for further analysis.
class DuplicateWorksheetWithPivot
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "SourceData";

        // Populate sample data for the pivot table
        sourceSheet.Cells["A1"].PutValue("Category");
        sourceSheet.Cells["B1"].PutValue("Amount");
        sourceSheet.Cells["A2"].PutValue("Fruit");
        sourceSheet.Cells["B2"].PutValue(100);
        sourceSheet.Cells["A3"].PutValue("Vegetable");
        sourceSheet.Cells["B3"].PutValue(200);
        sourceSheet.Cells["A4"].PutValue("Fruit");
        sourceSheet.Cells["B4"].PutValue(150);

        // Add a pivot table to the source worksheet
        int pivotIndex = sourceSheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
        PivotTable pivot = sourceSheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivot.CalculateData();

        // Duplicate the worksheet (including its pivot tables) using AddCopy
        int copiedIndex = workbook.Worksheets.AddCopy("SourceData");
        Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
        copiedSheet.Name = "CopiedData";

        // Refresh pivot tables in the copied worksheet to ensure they are up‑to‑date
        copiedSheet.RefreshPivotTables();

        // Save the workbook with the duplicated worksheet
        workbook.Save("DuplicatedWithPivot.xlsx");
    }
}
