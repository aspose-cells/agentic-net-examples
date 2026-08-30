// Title: Enable automatic refresh for an Aspose.Cells pivot table when the workbook is opened (C#)
// AI Prompts: Write C# code that creates an Excel workbook with sample data, adds a pivot table, and sets RefreshDataOnOpeningFile to true. | Modify an existing Aspose.Cells C# example to turn on automatic pivot table refresh on file load. | Show how to configure a pivot table's RefreshDataOnOpeningFile property in Aspose.Cells before saving the workbook.
// Common Searches: Aspose.Cells C# set pivot table to refresh on workbook open | How to enable RefreshDataOnOpeningFile for Excel pivot tables using Aspose.Cells .NET | Auto‑refresh pivot table when opening generated Excel file with Aspose.Cells | C# example for creating pivot table that updates automatically on file load
// Tags: Aspose.Cells pivot table refresh on open | RefreshDataOnOpeningFile property Aspose.Cells | C# create pivot table Excel workbook | programmatic pivot table refresh .NET | Excel workbook open pivot update Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotAutoRefresh
{
    // The program builds a new workbook, fills it with product and sales data, adds a pivot table, enables automatic refresh on opening by setting RefreshDataOnOpeningFile to true, and saves the file as PivotTable_AutoRefresh.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Sales");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(1000);
            dataSheet.Cells["A3"].PutValue("Banana");
            dataSheet.Cells["B3"].PutValue(2000);
            dataSheet.Cells["A4"].PutValue("Orange");
            dataSheet.Cells["B4"].PutValue(3000);

            // Add a pivot table based on the data range
            int pivotIndex = dataSheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
            PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];

            // Configure the pivot table (Product as row, Sales as data)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

            // Enable automatic refresh when the workbook is opened
            pivotTable.RefreshDataOnOpeningFile = true;

            // Save the workbook
            workbook.Save("PivotTable_AutoRefresh.xlsx");
        }
    }
}
