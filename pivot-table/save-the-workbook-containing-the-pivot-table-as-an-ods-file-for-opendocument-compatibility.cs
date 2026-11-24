// Title: Save an Excel workbook that includes a pivot table as an OpenDocument Spreadsheet (ODS) using Aspose.Cells for .NET
// Common Searches: how to export an Excel file with pivot tables to ODS format using Aspose.Cells C# | Aspose.Cells .NET save workbook as OpenDocument Spreadsheet preserving pivot tables | convert .xlsx containing pivot tables to .ods with Aspose.Cells example | C# code to save Excel workbook with pivot tables to ODS using SaveFormat.ODS | Aspose.Cells export pivot table data to OpenDocument spreadsheet in .NET application
// Tags: Aspose.Cells save workbook as ODS | export pivot table to OpenDocument format | convert Excel to ODS using C# | preserve pivot tables when saving as ODS | OpenDocument Spreadsheet export with Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an existing .xlsx workbook that contains a pivot table, optionally enumerates the pivot tables, and then saves the workbook as an OpenDocument Spreadsheet (ODS) file using Aspose.Cells' SaveFormat.ODS, ensuring the pivot table structure is retained.
class Program
{
    static void Main()
    {
        // Load the workbook that already contains a pivot table.
        // Adjust the file path as needed.
        Workbook workbook = new Workbook("input.xlsx");

        // (Optional) Verify that the workbook contains at least one pivot table.
        // This step is not mandatory for saving, but demonstrates access to pivot tables.
        // foreach (Worksheet sheet in workbook.Worksheets)
        // {
        //     foreach (PivotTable pt in sheet.PivotTables)
        //     {
        //         Console.WriteLine($"Pivot Table: {pt.Name} in sheet {sheet.Name}");
        //     }
        // }

        // Save the workbook in OpenDocument Spreadsheet (ODS) format for compatibility.
        workbook.Save("output.ods", SaveFormat.ODS);
    }
}
