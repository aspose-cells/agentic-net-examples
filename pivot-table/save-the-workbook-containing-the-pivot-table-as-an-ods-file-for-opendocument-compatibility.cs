// Title: Export a Pivot Table to ODS with Aspose.Cells in C#
// Description: Demonstrates how to create a workbook, add a pivot table, configure OdsSaveOptions (LibreOffice generator), and save the file as an OpenDocument Spreadsheet (ODS) using Aspose.Cells for .NET.
// Keywords: Aspose.Cells ODS export C# | pivot table to ODS | OdsSaveOptions LibreOffice | C# OpenDocument spreadsheet | save Excel pivot as ODS | Aspose.Cells .NET example
// Common Searches: how to save pivot table as ODS using Aspose.Cells | C# export Excel pivot to OpenDocument format | Aspose.Cells OdsSaveOptions example | convert Excel workbook with pivot to ODS in .NET | Aspose.Cells save workbook as LibreOffice ODS
// Developer Intent: Generate an OpenDocument Spreadsheet (ODS) that retains a pivot table created with Aspose.Cells.
// Use Cases: Provide cross‑platform sales reports for LibreOffice or OpenOffice users. | Automate creation of ODS files containing pivot analysis for email or web distribution. | Integrate ODS export into a .NET reporting pipeline that supports both Excel and OpenDocument formats.
// AI Prompts: Show C# code to add multiple data fields to a pivot table before exporting it to ODS with Aspose.Cells. | Explain how to set compression level and password protection in OdsSaveOptions while keeping the pivot table intact. | Provide steps to load an existing ODS file, modify its pivot table, and re‑save it using Aspose.Cells in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Ods;

// Demonstrates how to create a workbook, add a pivot table, configure OdsSaveOptions (LibreOffice generator), and save the file as an OpenDocument Spreadsheet (ODS) using Aspose.Cells for .NET.
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

        // Configure ODS save options (optional settings)
        OdsSaveOptions saveOptions = new OdsSaveOptions
        {
            GeneratorType = OdsGeneratorType.LibreOffice
        };

        // Save the workbook containing the pivot table as an ODS file
        workbook.Save("PivotTable.ods", saveOptions);
    }
}
