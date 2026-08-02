// Title: Get RefreshedByWho and RefreshDate of an Aspose.Cells PivotTable in C#
// Description: Demonstrates how to create a workbook, add a PivotTable, refresh it, and read the RefreshedByWho and RefreshDate properties with Aspose.Cells for .NET. The example also shows how to persist and retrieve this audit information after saving and reloading the file.
// Keywords: Aspose.Cells | PivotTable | RefreshedByWho | RefreshDate | C# | .NET | audit metadata | retrieve pivot refresh user | pivot table refresh timestamp | read pivot properties
// Common Searches: Aspose.Cells get pivot table refreshed by who | C# read RefreshDate of PivotTable | How to obtain pivot refresh user Aspose.Cells | Retrieve pivot table audit info .NET | Aspose.Cells PivotTable refresh metadata example
// Developer Intent: Extract the user name and timestamp of the last refresh operation for a PivotTable using Aspose.Cells.
// Use Cases: Display the refresh author and date in a reporting UI. | Store refresh metadata for compliance and audit trails. | Log pivot table refresh details when a workbook is opened or processed.
// AI Prompts: Write C# code that loads an existing workbook containing a PivotTable and prints its RefreshedByWho and RefreshDate values using Aspose.Cells. | Explain how Aspose.Cells sets the RefreshedByWho property during a PivotTable refresh and how to customize the value. | Provide a step‑by‑step guide to capture, store, and retrieve refresh user information for multiple PivotTables in one workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook, add a PivotTable, refresh it, and read the RefreshedByWho and RefreshDate properties with Aspose.Cells for .NET. The example also shows how to persist and retrieve this audit information after saving and reloading the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Fruit");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["B3"].PutValue(15);
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B4"].PutValue(8);

        // Add a pivot table to the worksheet
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "C3", "PivotTable1");
        PivotTable pt = sheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pt.AddFieldToArea(PivotFieldType.Row, 0);   // Fruit column as row field
        pt.AddFieldToArea(PivotFieldType.Data, 1);  // Quantity column as data field

        // Refresh the pivot table to populate the cache and set refresh metadata
        pt.RefreshData();
        pt.CalculateData();

        // Retrieve and display the user who last refreshed the pivot table
        Console.WriteLine("Refreshed By: " + pt.RefreshedByWho);
        Console.WriteLine("Refresh Date: " + pt.RefreshDate);

        // Save the workbook
        string filePath = "PivotRefreshInfo.xlsx";
        wb.Save(filePath);

        // Load the workbook again and read the same properties
        Workbook wb2 = new Workbook(filePath);
        PivotTable pt2 = wb2.Worksheets[0].PivotTables[0];
        Console.WriteLine("After reload - Refreshed By: " + pt2.RefreshedByWho);
        Console.WriteLine("After reload - Refresh Date: " + pt2.RefreshDate);
    }
}
