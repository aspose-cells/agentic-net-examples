// Title: Switch PivotTable to Outline Form with ShowInOutlineForm – Aspose.Cells C# Example
// Description: Shows how to create a workbook, populate sales data, build a pivot table, assign Date, Product, and Sales fields, enable the Outline layout via ShowInOutlineForm, refresh and calculate the pivot, and save the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | ShowInOutlineForm | PivotTable outline layout | C# | .NET | Excel pivot formatting | outline view | refresh pivot | calculate pivot | pivot table display mode
// Common Searches: Aspose.Cells ShowInOutlineForm example | C# set pivot table outline view | change pivot layout to outline Aspose.Cells | outline form pivot table .NET | switch pivot to outline using Aspose
// Developer Intent: Apply the Outline layout to a pivot table in an Excel workbook with Aspose.Cells for .NET.
// Use Cases: Generate a sales‑analysis workbook where the pivot hierarchy is displayed in outline mode for easier navigation. | Automate report pipelines that require existing pivot tables to be converted to outline layout before distribution. | Refresh and recalculate a pivot after changing its layout to ensure the displayed totals are up‑to‑date.
// AI Prompts: Write C# code that creates a pivot table with Aspose.Cells and switches it to outline form using ShowInOutlineForm. | Explain what ShowInOutlineForm does to a pivot table and compare it with other layout methods such as ShowInCompactForm and ShowInTabularForm. | Provide step‑by‑step instructions to refresh, calculate, and save a pivot table after applying the outline layout in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotOutlineDemo
{
    // Shows how to create a workbook, populate sales data, build a pivot table, assign Date, Product, and Sales fields, enable the Outline layout via ShowInOutlineForm, refresh and calculate the pivot, and save the result as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a worksheet that will hold the source data
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data
            dataSheet.Cells["A1"].PutValue("Date");
            dataSheet.Cells["B1"].PutValue("Product");
            dataSheet.Cells["C1"].PutValue("Sales");

            dataSheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            dataSheet.Cells["B2"].PutValue("Apple");
            dataSheet.Cells["C2"].PutValue(1000);

            dataSheet.Cells["A3"].PutValue(new DateTime(2023, 1, 2));
            dataSheet.Cells["B3"].PutValue("Banana");
            dataSheet.Cells["C3"].PutValue(2000);

            dataSheet.Cells["A4"].PutValue(new DateTime(2023, 1, 3));
            dataSheet.Cells["B4"].PutValue("Apple");
            dataSheet.Cells["C4"].PutValue(1500);

            // Add a worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table and obtain its reference
            int pivotIndex = pivotSheet.PivotTables.Add(
                "=Data!A1:C4",   // source data range
                "A3",            // top‑left cell of the pivot table
                "PivotTable1");  // pivot table name

            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add fields to the pivot table areas
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Switch the pivot table layout to Outline form
            pivotTable.ShowInOutlineForm();

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableOutlineFormDemo.xlsx");
        }
    }
}
