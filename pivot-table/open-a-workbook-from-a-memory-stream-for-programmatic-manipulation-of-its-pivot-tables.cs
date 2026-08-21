// Title: Load an Excel workbook from a MemoryStream and edit its PivotTable with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, save it to a MemoryStream, reload it, modify the first PivotTable (disable automatic refresh, refresh data, recalculate), and save the result to a file using Aspose.Cells for C#.
// Keywords: Aspose.Cells MemoryStream | C# load workbook from stream | modify PivotTable programmatically | RefreshDataOnOpeningFile | PivotTable RefreshData | PivotTable CalculateData | save workbook after stream manipulation | Excel pivot table .NET
// Common Searches: Aspose.Cells load workbook from MemoryStream | C# edit PivotTable after loading from stream | disable pivot refresh on opening file Aspose.Cells | refresh and calculate pivot data programmatically | save modified Excel file after stream processing
// Developer Intent: Load a workbook from a MemoryStream, change PivotTable settings, refresh its data, and write the updated file.
// Use Cases: Transfer an Excel file between services via a MemoryStream, adjust PivotTable behavior, and persist the changes. | Prevent automatic PivotTable refresh for large workbooks to improve load performance. | Programmatically recalculate PivotTable data after modifying the source range in memory.
// AI Prompts: Generate C# code that opens an Excel workbook from a MemoryStream, disables the PivotTable's RefreshDataOnOpeningFile, calls RefreshData and CalculateData, and saves the file. | Show how to reset a MemoryStream position before loading a workbook with Aspose.Cells and then update a PivotTable. | Explain why disabling automatic pivot refresh is useful for large Excel files when using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook, save it to a MemoryStream, reload it, modify the first PivotTable (disable automatic refresh, refresh data, recalculate), and save the result to a file using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a workbook with sample data and a pivot table.
        // ------------------------------------------------------------
        Workbook sourceWb = new Workbook();
        Worksheet ws = sourceWb.Worksheets[0];

        ws.Cells["A1"].PutValue("Category");
        ws.Cells["B1"].PutValue("Amount");
        ws.Cells["A2"].PutValue("Food");
        ws.Cells["B2"].PutValue(100);
        ws.Cells["A3"].PutValue("Drink");
        ws.Cells["B3"].PutValue(150);
        ws.Cells["A4"].PutValue("Food");
        ws.Cells["B4"].PutValue(200);

        // Add a pivot table based on the data range.
        int ptIndex = ws.PivotTables.Add("A1:B4", "D1", "SalesPivot");
        PivotTable pt = ws.PivotTables[ptIndex];
        pt.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pt.AddFieldToArea(PivotFieldType.Data, 1);  // Amount as data field

        // ------------------------------------------------------------
        // 2. Save the workbook into a MemoryStream.
        // ------------------------------------------------------------
        using (MemoryStream ms = new MemoryStream())
        {
            sourceWb.Save(ms, SaveFormat.Xlsx);
            ms.Position = 0; // Reset stream position for reading.

            // ------------------------------------------------------------
            // 3. Load the workbook from the MemoryStream.
            // ------------------------------------------------------------
            Workbook wb = new Workbook(ms);

            // ------------------------------------------------------------
            // 4. Manipulate the pivot table in the loaded workbook.
            // ------------------------------------------------------------
            Worksheet loadedWs = wb.Worksheets[0];
            PivotTable loadedPt = loadedWs.PivotTables[0];

            // Example: prevent automatic refresh when the file is opened.
            loadedPt.RefreshDataOnOpeningFile = false;

            // Refresh and recalculate the pivot table to reflect any data changes.
            loadedPt.RefreshData();
            loadedPt.CalculateData();

            // ------------------------------------------------------------
            // 5. Save the modified workbook to a physical file.
            // ------------------------------------------------------------
            wb.Save("ModifiedFromStream.xlsx", SaveFormat.Xlsx);
        }
    }
}
