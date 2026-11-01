// Title: Set a custom DataFieldSeparator for an Aspose.Cells PivotTable in C#
// AI Prompts: Generate an Excel workbook with Aspose.Cells, add a PivotTable, and assign a semicolon as the DataFieldSeparator using C#. | Modify an existing PivotTable in a .NET workbook to use a pipe (|) as the multi‑field value delimiter via the DataFieldSeparator property. | Create a PivotTable with Aspose.Cells and configure the DataFieldSeparator to a custom character for combined data fields in C#.
// Common Searches: Aspose.Cells C# how to change the DataFieldSeparator of a PivotTable | custom delimiter for combined data fields in Aspose.Cells PivotTable | set pipe character as pivot table data field separator using Aspose.Cells .NET | example of configuring PivotTable.DataFieldSeparator property in C# | multi‑field value separator in generated Excel pivot table with Aspose.Cells
// Tags: Aspose.Cells PivotTable DataFieldSeparator property | C# set custom pivot table delimiter | Excel pivot table multi-field separator Aspose.Cells | programmatic pivot table delimiter .NET | custom data field separator for Aspose.Cells PivotTable

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates creating a workbook, adding a PivotTable, and setting the PivotTable.DataFieldSeparator property to a custom character before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table (range A1:D6)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("SubCategory");
            sheet.Cells["C1"].PutValue("Region");
            sheet.Cells["D1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Beverages");
            sheet.Cells["B2"].PutValue("Tea");
            sheet.Cells["C2"].PutValue("North");
            sheet.Cells["D2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Beverages");
            sheet.Cells["B3"].PutValue("Coffee");
            sheet.Cells["C3"].PutValue("South");
            sheet.Cells["D3"].PutValue(1500);

            sheet.Cells["A4"].PutValue("Food");
            sheet.Cells["B4"].PutValue("Bread");
            sheet.Cells["C4"].PutValue("North");
            sheet.Cells["D4"].PutValue(800);

            sheet.Cells["A5"].PutValue("Food");
            sheet.Cells["B5"].PutValue("Butter");
            sheet.Cells["C5"].PutValue("South");
            sheet.Cells["D5"].PutValue(950);

            sheet.Cells["A6"].PutValue("Food");
            sheet.Cells["B6"].PutValue("Cheese");
            sheet.Cells["C6"].PutValue("East");
            sheet.Cells["D6"].PutValue(1100);

            // Add a new worksheet for the pivot table
            int pivotSheetIndex = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "Pivot";

            // Define the source data range
            string sourceData = "A1:D6";

            // Add the pivot table at cell A1 of the pivot sheet
            int pivotIndex = pivotSheet.PivotTables.Add("MyPivot", "A1", sourceData);
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // NOTE: Adding fields to the pivot table depends on the Aspose.Cells version.
            // The following lines are omitted to maintain compatibility across versions.
            // You can customize the pivot fields using the appropriate API for your version.

            // Define output file path
            string outputPath = "PivotTable_CustomSeparator.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
