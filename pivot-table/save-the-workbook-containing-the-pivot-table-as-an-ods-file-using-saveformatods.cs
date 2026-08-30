// Title: Generate a pivot table from sample data and export the workbook to OpenDocument Spreadsheet (ODS) using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, fills it with sample data, builds a pivot table, and saves the file as ODS using Aspose.Cells. | Show how to add row and data fields to an Aspose.Cells pivot table before exporting the workbook to OpenDocument format. | Demonstrate using Workbook.Save with SaveFormat.ODS to write a workbook containing a pivot table to disk in C#.
// Common Searches: Aspose.Cells C# create pivot table and save as .ods | How to export a workbook with a pivot table to OpenDocument Spreadsheet using Aspose.Cells | SaveFormat.ODS example for pivot tables in .NET | C# Aspose.Cells pivot table ODS export sample code | OpenDocument Spreadsheet output from Aspose.Cells pivot table
// Tags: Aspose.Cells create pivot table | Aspose.Cells save workbook as ODS | pivot table row field configuration Aspose.Cells | export workbook to OpenDocument format C# | SaveFormat.ODS usage with pivot tables

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotOdsDemo
{
    // The program creates a new workbook, populates it with sample data, adds a pivot table with row and data fields, and saves the workbook as an OpenDocument Spreadsheet (ODS) file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Sales";
            cells["A2"].Value = "Apple";
            cells["B2"].Value = 1200;
            cells["A3"].Value = "Orange";
            cells["B3"].Value = 850;
            cells["A4"].Value = "Banana";
            cells["B4"].Value = 430;

            // Add a pivot table based on the data range A1:B4, place it at E5
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E5", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Product as row field, Sales as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Product
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Sales

            // Save the workbook as ODS using the SaveFormat.ODS enum
            workbook.Save("SalesPivotTable.ods", SaveFormat.ODS);
        }
    }
}
