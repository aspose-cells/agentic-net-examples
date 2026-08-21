// Title: C# – Retrieve PivotTable RefreshedByWho and RefreshDate using Aspose.Cells
// Description: Demonstrates how to create a workbook, add sample data, build a PivotTable, refresh it, and read the RefreshedByWho and RefreshDate properties. The example also saves the file, reloads it, and shows that the refresh metadata persists.
// Keywords: Aspose.Cells PivotTable RefreshedByWho | C# get pivot refresh user | Aspose.Cells RefreshDate property | read pivot refresh metadata .NET | pivot table audit trail Aspose | Aspose.Cells PivotTable example
// Common Searches: How to get the user who refreshed a PivotTable with Aspose.Cells | Aspose.Cells RefreshedByWho property C# example | Read PivotTable refresh date after saving Aspose.Cells | Retrieve pivot refresh information from saved workbook | Aspose.Cells PivotTable audit information
// Developer Intent: The developer needs to obtain the RefreshedByWho value (and optionally the RefreshDate) of a PivotTable to identify who last refreshed it.
// Use Cases: Log the username and timestamp each time a PivotTable is refreshed for compliance reporting. | Show end‑users the last refresh details in a dashboard after opening a workbook. | Validate that a PivotTable was refreshed by an authorized account before running further calculations.
// AI Prompts: Generate C# code that sets RefreshedByWho manually before saving a PivotTable with Aspose.Cells. | Provide a comparison script that checks RefreshedByWho values across two workbooks. | Explain how Aspose.Cells populates RefreshedByWho and RefreshDate when RefreshData is called.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshInfo
{
    // Demonstrates how to create a workbook, add sample data, build a PivotTable, refresh it, and read the RefreshedByWho and RefreshDate properties. The example also saves the file, reloads it, and shows that the refresh metadata persists.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["A2"].Value = "Food";
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["A3"].Value = "Drink";
            sheet.Cells["B3"].Value = 80;
            sheet.Cells["A4"].Value = "Food";
            sheet.Cells["B4"].Value = 150;
            sheet.Cells["A5"].Value = "Drink";
            sheet.Cells["B5"].Value = 70;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot fields
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivot.AddFieldToArea(PivotFieldType.Data, 1); // Amount as data field

            // Refresh the pivot table so that RefreshByWho gets populated
            pivot.RefreshData();
            pivot.CalculateData();

            // Output the user who last refreshed the pivot table
            Console.WriteLine("Pivot Table refreshed by: " + pivot.RefreshedByWho);
            Console.WriteLine("Refresh date: " + pivot.RefreshDate);

            // Save the workbook to demonstrate persistence of the property
            string filePath = "PivotRefreshInfoDemo.xlsx";
            workbook.Save(filePath);

            // Reload the workbook and read the property again
            Workbook loadedWb = new Workbook(filePath);
            PivotTable loadedPivot = loadedWb.Worksheets[0].PivotTables[0];
            Console.WriteLine("After reload - refreshed by: " + loadedPivot.RefreshedByWho);
            Console.WriteLine("After reload - refresh date: " + loadedPivot.RefreshDate);
        }
    }
}
