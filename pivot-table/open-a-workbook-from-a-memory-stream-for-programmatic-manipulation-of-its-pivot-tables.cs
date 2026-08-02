// Title: Load an Excel workbook from a MemoryStream, modify its pivot table, and save with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook with a pivot table, save it to a MemoryStream, reload it, disable automatic refresh, manually refresh the pivot data, and write the updated file back to a new stream using Aspose.Cells for C#.
// Keywords: Aspose.Cells load workbook from memory stream | C# pivot table manipulation | refresh pivot data programmatically | disable pivot auto refresh Aspose | save workbook to stream .NET | memory stream Excel example | Aspose.Cells pivot table API
// Common Searches: Aspose.Cells open Excel from MemoryStream | how to refresh pivot table after loading from stream | save modified workbook to MemoryStream C# | disable RefreshDataOnOpeningFile Aspose.Cells | pivot table programmatic update Aspose
// Developer Intent: Read an Excel file from a MemoryStream, adjust pivot‑table settings, refresh its data, and output the result to another stream without touching the file system.
// Use Cases: Web API endpoint that receives an uploaded XLSX as a byte array, updates pivot tables in‑memory, and returns the modified file as a response stream. | Batch processing of dozens of workbooks stored in a message queue, where each file is edited in memory to avoid disk I/O and then forwarded to downstream services. | Server‑side automation that generates a report, adds a pivot table, saves the workbook to a MemoryStream, and streams it directly to a client browser.
// AI Prompts: Write C# code using Aspose.Cells to load an Excel workbook from a MemoryStream, turn off RefreshDataOnOpeningFile for all pivot tables, call RefreshData, and save the workbook to a new MemoryStream. | Explain the steps required to programmatically refresh a pivot table after loading a workbook from a stream with Aspose.Cells, including any necessary property configurations. | Provide an ASP.NET Core controller example that accepts an Excel file as a byte array, updates its pivot table fields with Aspose.Cells, and returns the edited file as a byte array.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook with a pivot table, save it to a MemoryStream, reload it, disable automatic refresh, manually refresh the pivot data, and write the updated file back to a new stream using Aspose.Cells for C#.
public class PivotTableMemoryStreamDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // ------------------------------------------------------------
        // 1. Create a sample workbook that contains a pivot table.
        // ------------------------------------------------------------
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Populate source data.
        sourceSheet.Cells["A1"].PutValue("Category");
        sourceSheet.Cells["B1"].PutValue("Amount");
        sourceSheet.Cells["A2"].PutValue("Food");
        sourceSheet.Cells["B2"].PutValue(100);
        sourceSheet.Cells["A3"].PutValue("Drink");
        sourceSheet.Cells["B3"].PutValue(150);
        sourceSheet.Cells["A4"].PutValue("Food");
        sourceSheet.Cells["B4"].PutValue(200);

        // Add a pivot table based on the data range.
        int pivotIndex = sourceSheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
        PivotTable pivot = sourceSheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Amount as data field

        // ------------------------------------------------------------
        // 2. Save the workbook into a MemoryStream (XLSX format).
        // ------------------------------------------------------------
        using (MemoryStream memoryStream = new MemoryStream())
        {
            sourceWorkbook.Save(memoryStream, SaveFormat.Xlsx);
            memoryStream.Position = 0; // Reset stream position for reading.

            // ------------------------------------------------------------
            // 3. Load the workbook from the MemoryStream.
            // ------------------------------------------------------------
            Workbook loadedWorkbook = new Workbook(memoryStream);

            // ------------------------------------------------------------
            // 4. Manipulate pivot tables programmatically.
            // ------------------------------------------------------------
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            if (loadedSheet.PivotTables.Count > 0)
            {
                PivotTable loadedPivot = loadedSheet.PivotTables[0];

                // Disable automatic refresh on opening.
                loadedPivot.RefreshDataOnOpeningFile = false;

                // Manually refresh the pivot data.
                loadedPivot.RefreshData();
            }

            // ------------------------------------------------------------
            // 5. Save the modified workbook back to a new MemoryStream.
            // ------------------------------------------------------------
            using (MemoryStream outStream = loadedWorkbook.SaveToStream())
            {
                string outputPath = "ModifiedFromStream.xlsx";
                File.WriteAllBytes(outputPath, outStream.ToArray());
                Console.WriteLine($"Modified workbook saved to {outputPath}");
            }
        }
    }
}
