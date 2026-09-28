// Title: Enable ShowErrorValues in an Aspose.Cells PivotTable using C# to display calculation errors
// AI Prompts: Write C# code that creates a workbook with a pivot table and activates the property that displays error values directly in the pivot cells using Aspose.Cells. | Provide a C# example that configures a pivot table to reveal original error values instead of a placeholder, then recalculates the pivot data. | Generate a C# snippet that sets a custom error string for a pivot table, clears it to show the actual error, and saves the workbook.
// Common Searches: Aspose.Cells C# how to display calculation errors in pivot table cells | Enable error value visibility for pivot tables with Aspose.Cells .NET | Show original error messages in Aspose.Cells pivot table output | Set pivot table to show #DIV/0! errors using Aspose.Cells C#
// Tags: Aspose.Cells ShowErrorValues API | C# pivot table error display | Aspose.Cells error message configuration | PivotTable calculation error display .NET | Aspose.Cells workbook pivot error handling

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotShowErrorValuesDemo
{
    // The example creates a new workbook, populates it with sample product and sales data, adds a pivot table, assigns the Product column as a row field and the Sales column as a data field, enables the ShowErrorValues (DisplayErrorString) flag to show calculation errors, clears any custom error string to display the original error, recalculates the pivot data, and saves the workbook as PivotTableShowErrorValuesDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Sales";
            cells["A2"].Value = "Bike";
            cells["B2"].Value = 1000;
            cells["A3"].Value = "Car";
            cells["B3"].Value = 2000;
            cells["A4"].Value = "Bike";
            cells["B4"].Value = 1500;
            cells["A5"].Value = "Car";
            cells["B5"].Value = 2500;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table (Product as row, Sales as data)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

            // Enable display of calculation errors directly within pivot cells
            pivotTable.DisplayErrorString = true;
            // Optionally set a custom error string (empty string will show the original error)
            pivotTable.ErrorString = "";

            // Calculate the pivot data so that any errors are reflected
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableShowErrorValuesDemo.xlsx");
        }
    }
}
