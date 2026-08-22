// Title: Load an Excel workbook from a MemoryStream, refresh its pivot tables, and disable auto‑refresh using Aspose.Cells for .NET
// AI Prompts: Open a workbook from a MemoryStream, call RefreshPivotTables on the worksheet, set RefreshDataOnOpeningFile to false for each pivot table, then save the workbook to a new stream. | Read an in‑memory Excel file, modify pivot table settings (e.g., turn off automatic refresh), and write the updated workbook as an Xlsx stream with Aspose.Cells.
// Common Searches: Aspose.Cells .NET load workbook from MemoryStream and update pivot tables | How to refresh all pivot tables after opening an Excel file from a stream using Aspose.Cells | Disable pivot table refresh on file open with Aspose.Cells C# | Convert a workbook saved to a MemoryStream (xls) to Xlsx after modifying pivot tables | Programmatically manipulate pivot tables after loading Excel from a stream in C#
// Tags: memory stream workbook loading Aspose.Cells | refresh all pivot tables Aspose.Cells | disable pivot table auto refresh Aspose.Cells | convert workbook stream to Xlsx Aspose.Cells | programmatic pivot table manipulation .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example demonstrates how to create a workbook with a pivot table, save it to a MemoryStream, reload it from that stream, refresh all pivot tables, turn off automatic refresh on opening, and finally save the modified workbook as an Xlsx file.
class Program
{
    static void Main()
    {
        // 1. Create a workbook with sample data and a pivot table
        Workbook sourceWb = new Workbook();
        Worksheet ws = sourceWb.Worksheets[0];
        Cells cells = ws.Cells;

        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Amount");
        cells["A2"].PutValue("Food");
        cells["B2"].PutValue(100);
        cells["A3"].PutValue("Drink");
        cells["B3"].PutValue(150);
        cells["A4"].PutValue("Food");
        cells["B4"].PutValue(200);

        // Add a pivot table based on the data range
        int ptIndex = ws.PivotTables.Add("A1:B4", "D1", "Pivot1");
        PivotTable pt = ws.PivotTables[ptIndex];
        pt.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pt.AddFieldToArea(PivotFieldType.Data, 1);  // Amount as data field

        // 2. Save the workbook to a memory stream (xls format) using the provided SaveToStream method
        MemoryStream memStream = sourceWb.SaveToStream();

        // Reset the stream position before reading
        memStream.Position = 0;

        // 3. Load the workbook from the memory stream using the Stream constructor
        Workbook wb = new Workbook(memStream);

        // 4. Manipulate pivot tables programmatically
        Worksheet loadedWs = wb.Worksheets[0];

        // Refresh all pivot tables in the worksheet
        loadedWs.RefreshPivotTables();

        // Example: disable automatic refresh when the file is opened
        if (loadedWs.PivotTables.Count > 0)
        {
            PivotTable loadedPt = loadedWs.PivotTables[0];
            loadedPt.RefreshDataOnOpeningFile = false;
        }

        // 5. Save the modified workbook back to a new memory stream (xlsx format)
        using (MemoryStream outStream = new MemoryStream())
        {
            wb.Save(outStream, SaveFormat.Xlsx);
            // For demonstration, write the result to a physical file
            File.WriteAllBytes("ModifiedWorkbook.xlsx", outStream.ToArray());
        }

        // Clean up resources
        memStream.Dispose();
    }
}
