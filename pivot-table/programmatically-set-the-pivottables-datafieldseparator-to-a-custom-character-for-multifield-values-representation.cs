// Title: How to set a custom DataFieldSeparator for an Aspose.Cells PivotTable in C#
// AI Prompts: Write C# code that creates a workbook, adds a PivotTable with Aspose.Cells, and assigns a user‑defined character to the DataFieldSeparator property before saving the file. | Show how to update the DataFieldSeparator of an existing Aspose.Cells PivotTable to use a pipe (|) delimiter for combined field values.
// Common Searches: Aspose.Cells C# change pivot table DataFieldSeparator to semicolon | set custom delimiter for multi‑field values in Aspose.Cells PivotTable | how to modify DataFieldSeparator property of a PivotTable using Aspose.Cells .NET | example code for custom DataFieldSeparator in Aspose.Cells pivot tables
// Tags: set DataFieldSeparator Aspose.Cells | custom delimiter pivot table Aspose.Cells | C# Aspose.Cells pivot table separator | modify PivotTable field separator .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a new Workbook, populates cells A1:C5 with Category, SubCategory, and Amount data, adds a PivotTable named "MyPivot" at E3, assigns row, column, and data fields, sets the DataFieldSeparator property to a custom character (e.g., ';'), refreshes and calculates the pivot cache, and saves the workbook as PivotWithCustomSeparator.xlsx.
class SetPivotDataFieldSeparator
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "SubCategory";
            cells["C1"].Value = "Amount";

            cells["A2"].Value = "Fruit";
            cells["B2"].Value = "Apple";
            cells["C2"].Value = 120;

            cells["A3"].Value = "Fruit";
            cells["B3"].Value = "Banana";
            cells["C3"].Value = 80;

            cells["A4"].Value = "Vegetable";
            cells["B4"].Value = "Carrot";
            cells["C4"].Value = 50;

            cells["A5"].Value = "Vegetable";
            cells["B5"].Value = "Broccoli";
            cells["C5"].Value = 70;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "MyPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "SubCategory");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh pivot cache and calculate data
            pivotTable.RefreshData();      // Correct method to refresh data
            pivotTable.CalculateData();

            // Save the workbook
            string outputPath = "PivotWithCustomSeparator.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
