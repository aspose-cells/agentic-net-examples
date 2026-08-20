// Title: Export a Pivot Table to ODS with Aspose.Cells (C#)
// Description: Creates a workbook, adds a pivot table on sample product‑sales data, and saves the file as an OpenDocument Spreadsheet (ODS) using Aspose.Cells SaveFormat.Ods.
// Keywords: Aspose.Cells | C# | pivot table | ODS export | SaveFormat.Ods | OpenDocument Spreadsheet | LibreOffice compatibility
// Common Searches: Aspose.Cells save pivot table as ODS | C# export pivot table to OpenDocument Spreadsheet | How to use SaveFormat.Ods with a pivot table | Generate ODS file from Aspose.Cells .NET | OpenDocument spreadsheet pivot export example
// Developer Intent: Export a workbook that contains a pivot table to an ODS file.
// Use Cases: Produce a sales summary with a pivot table and share it as an ODS file for LibreOffice users. | Automate generation of cross‑platform spreadsheets that include pivot analyses. | Integrate ODS pivot export into a .NET web service that returns downloadable reports.
// AI Prompts: Write C# code that builds a workbook, adds a pivot table, and saves it as ODS with Aspose.Cells, including error handling. | Explain how to adjust pivot table settings before exporting to ODS using Aspose.Cells for .NET. | Show how to apply ODS‑specific save options when exporting a workbook containing a pivot table.

using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook, adds a pivot table on sample product‑sales data, and saves the file as an OpenDocument Spreadsheet (ODS) using Aspose.Cells SaveFormat.Ods.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(1000);
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["B3"].PutValue(2000);
        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B4"].PutValue(3000);

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "E5", "SalesPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field

        // Save the workbook as an ODS file using SaveFormat.Ods
        workbook.Save("SalesPivot.ods", SaveFormat.Ods);
    }
}
