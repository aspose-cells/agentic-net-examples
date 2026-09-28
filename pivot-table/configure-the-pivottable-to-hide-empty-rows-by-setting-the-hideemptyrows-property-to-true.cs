// Title: C# example: hide empty rows in an Aspose.Cells PivotTable by setting ShowEmptyRow to false
// AI Prompts: Write C# code that creates an Excel workbook, adds sample data with blank rows, builds a pivot table, and disables ShowEmptyRow to exclude those rows using Aspose.Cells. | Modify an existing Aspose.Cells PivotTable in C# to prevent rows without data from appearing by setting the ShowEmptyRow property to false and then recalculate the pivot. | Generate a complete C# snippet that demonstrates configuring a pivot table to hide empty rows and saves the result as an .xlsx file with Aspose.Cells.
// Common Searches: Aspose.Cells C# hide empty rows in pivot table | ShowEmptyRow property false example Aspose.Cells .NET | How to remove blank rows from a pivot table using Aspose.Cells in C# | C# Aspose.Cells pivot table exclude rows with no data | Configure Aspose.Cells pivot table to not display empty rows
// Tags: Aspose.Cells pivot table hide empty rows | C# ShowEmptyRow property Aspose.Cells | Aspose.Cells pivot table row visibility control | Excel pivot table blank row exclusion C# | Aspose.Cells calculate pivot data

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a workbook, fills it with data that includes empty rows, adds a pivot table on range A1:C5, assigns Category and Product as row fields and Sales as a data field, sets ShowEmptyRow to false to hide rows lacking data, recalculates the pivot, and saves the file as PivotTableHideEmptyRows.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with some empty rows
        sheet.Cells["A1"].Value = "Category";
        sheet.Cells["B1"].Value = "Product";
        sheet.Cells["C1"].Value = "Sales";

        sheet.Cells["A2"].Value = "Electronics";
        sheet.Cells["B2"].Value = "TV";
        sheet.Cells["C2"].Value = 1000;

        sheet.Cells["A3"].Value = "Electronics";
        // Empty row (no product, no sales)
        sheet.Cells["B3"].Value = "";
        sheet.Cells["C3"].Value = "";

        sheet.Cells["A4"].Value = "Furniture";
        sheet.Cells["B4"].Value = "Chair";
        sheet.Cells["C4"].Value = 500;

        sheet.Cells["A5"].Value = "Furniture";
        // Empty row (no product, no sales)
        sheet.Cells["B5"].Value = "";
        sheet.Cells["C5"].Value = "";

        // Add a pivot table based on the data range
        PivotTableCollection pivotTables = sheet.PivotTables;
        int pivotIndex = pivotTables.Add("A1:C5", "E3", "PivotTable1");
        PivotTable pivotTable = pivotTables[pivotIndex];

        // Add fields to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Hide empty rows in the pivot table
        // Setting ShowEmptyRow to false excludes empty rows (hides them)
        pivotTable.ShowEmptyRow = false;

        // Calculate the pivot table data
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("PivotTableHideEmptyRows.xlsx");
    }
}
