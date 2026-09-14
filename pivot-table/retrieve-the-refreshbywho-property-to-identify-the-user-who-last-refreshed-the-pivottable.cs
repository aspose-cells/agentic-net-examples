// Title: How to read the RefreshedByWho and RefreshDate properties of a PivotTable with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds a pivot table, calls RefreshData and CalculateData, then prints pivotTable.RefreshedByWho and pivotTable.RefreshDate. | Show how to access pivot table refresh metadata (user name and timestamp) using the Aspose.Cells PivotTable API in a .NET console application. | Generate a minimal Aspose.Cells example that demonstrates retrieving the last refreshed user and date from a PivotTable after refreshing it.
// Common Searches: Aspose.Cells C# get the user who refreshed a pivot table | How to obtain RefreshDate from a PivotTable using Aspose.Cells .NET | C# example for reading RefreshedByWho property of Aspose.Cells PivotTable | Retrieve pivot table refresh metadata with Aspose.Cells for .NET | Aspose.Cells PivotTable RefreshData and RefreshDate usage in C#
// Tags: Aspose.Cells PivotTable refresh metadata | C# read RefreshedByWho Aspose.Cells | Aspose.Cells get pivot refresh date | PivotTable RefreshData usage Aspose.Cells | Aspose.Cells .NET pivot user info

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a workbook, populates sample data, adds a pivot table, refreshes it, and then outputs the RefreshedByWho and RefreshDate properties to show who last refreshed the pivot and when, before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Fruit");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["B3"].PutValue(15);
        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B4"].PutValue(8);

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "FruitPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Configure the pivot table (row field and data field)
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Fruit column
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Quantity column

        // Refresh the pivot table to populate RefreshedByWho and RefreshDate
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Retrieve and display the user who last refreshed the pivot table
        Console.WriteLine("Refreshed By: " + pivotTable.RefreshedByWho);
        Console.WriteLine("Refresh Date: " + pivotTable.RefreshDate);

        // Save the workbook (optional, demonstrates persistence)
        string outputPath = "PivotRefreshInfo.xlsx";
        workbook.Save(outputPath);
    }
}
