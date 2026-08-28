// Title: Enable DetailLink in Aspose.Cells for .NET to add hyperlinks from pivot table master rows to generated detail worksheets (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to build a pivot table, invoke ShowDetail with newSheet=true, and insert a hyperlink from the master cell to the A1 cell of the automatically created detail worksheet. | Write a reusable C# method that, given a Workbook and a pivot table, creates a detail sheet via ShowDetail and adds a DetailLink hyperlink from a specified master row cell to the new sheet using Aspose.Cells.
// Common Searches: Aspose.Cells C# enable DetailLink to hyperlink pivot row to detail sheet | How to add a hyperlink from a master pivot table row to a generated detail worksheet using Aspose.Cells | ShowDetail new worksheet hyperlink example Aspose.Cells .NET | Create master‑detail Excel links with Aspose.Cells pivot tables in C#
// Tags: Aspose.Cells pivot ShowDetail hyperlink | C# Aspose.Cells master‑detail worksheet linking | Aspose.Cells create detail worksheet hyperlink | Aspose.Cells hyperlink to generated sheet | Aspose.Cells pivot table detail link C#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example demonstrates how to enable the DetailLink feature in Aspose.Cells for .NET: it creates a workbook with a master sheet, builds a pivot table, calls ShowDetail to generate a separate detail worksheet, and adds a hyperlink from the master pivot cell to cell A1 of the newly created detail sheet, then saves the file.
class DetailLinkDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet (master sheet)
            Workbook workbook = new Workbook();
            Worksheet masterSheet = workbook.Worksheets[0];
            masterSheet.Name = "Master";

            // Populate master sheet with sample data for the pivot table
            masterSheet.Cells["A1"].PutValue("Category");
            masterSheet.Cells["B1"].PutValue("Amount");
            masterSheet.Cells["A2"].PutValue("Food");
            masterSheet.Cells["B2"].PutValue(120);
            masterSheet.Cells["A3"].PutValue("Food");
            masterSheet.Cells["B3"].PutValue(80);
            masterSheet.Cells["A4"].PutValue("Travel");
            masterSheet.Cells["B4"].PutValue(200);
            masterSheet.Cells["A5"].PutValue("Travel");
            masterSheet.Cells["B5"].PutValue(150);

            // Add a pivot table on the master sheet
            int pivotIndex = masterSheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivot = masterSheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh pivot cache data (correct API) and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Show detail for the first data item (rowOffset = 0, columnOffset = 0)
            // newSheet = true creates a new worksheet that contains the detail data
            pivot.ShowDetail(0, 0, true, 0, 0);

            // The detail worksheet is the last worksheet added to the workbook
            Worksheet detailSheet = workbook.Worksheets[workbook.Worksheets.Count - 1];
            detailSheet.Name = "Detail";

            // Add a hyperlink on the master sheet that points to the detail sheet
            // Link cell D3 (where the pivot table starts) to cell A1 of the detail sheet
            masterSheet.Hyperlinks.Add("D3", 1, 1, $"'{detailSheet.Name}'!A1");

            // Save the workbook
            workbook.Save("DetailLinkDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
