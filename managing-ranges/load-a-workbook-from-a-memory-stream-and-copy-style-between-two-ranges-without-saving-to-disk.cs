// Title: Copy cell formatting between two equal-sized ranges after loading a workbook from a memory stream with Aspose.Cells for .NET
// AI Prompts: Load an Excel workbook from a byte array using a MemoryStream and transfer the style from range A1:B2 to range D5:E6 without writing the file to disk. | Validate that the source and destination ranges have identical dimensions, then copy each cell's style from the source range to the destination range entirely in memory using Aspose.Cells.
// Common Searches: asp.net copy cell style between ranges using Aspose.Cells memory stream | load Excel workbook from byte array and duplicate formatting in Aspose.Cells C# | how to transfer range formatting without saving workbook Aspose.Cells .NET | in‑memory range style copy Aspose.Cells example
// Tags: copy cell style Aspose.Cells | load workbook from memory stream C# | in-memory range formatting Aspose.Cells | validate equal range dimensions Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file from a byte array via a MemoryStream, creates matching source and destination ranges, checks that their sizes match, and copies each cell's style from the source range to the destination range, all without persisting the workbook to disk.
class Program
{
    static void Main()
    {
        try
        {
            // Obtain the Excel file as a byte array (could come from DB, network, etc.)
            byte[] workbookData = GetWorkbookBytes();

            // Load the workbook from a memory stream (no disk I/O)
            using (MemoryStream ms = new MemoryStream(workbookData))
            {
                Workbook wb = new Workbook(ms);

                // Access the first worksheet
                Worksheet ws = wb.Worksheets[0];

                // Define source and destination ranges
                // Example: copy style from A1:B2 to D5:E6
                Aspose.Cells.Range srcRange = ws.Cells.CreateRange("A1", "B2");
                Aspose.Cells.Range destRange = ws.Cells.CreateRange("D5", "E6");

                // Ensure both ranges have the same dimensions
                if (srcRange.RowCount != destRange.RowCount ||
                    srcRange.ColumnCount != destRange.ColumnCount)
                {
                    throw new InvalidOperationException("Source and destination ranges must be the same size.");
                }

                // Copy style cell by cell
                for (int i = 0; i < srcRange.RowCount; i++)
                {
                    for (int j = 0; j < srcRange.ColumnCount; j++)
                    {
                        Cell srcCell = srcRange[i, j];
                        Cell destCell = destRange[i, j];
                        destCell.SetStyle(srcCell.GetStyle());
                    }
                }

                // Workbook remains in memory; no saving to disk required
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to provide a sample workbook as a byte array.
    // Replace with actual source of the Excel file in real usage.
    static byte[] GetWorkbookBytes()
    {
        Workbook tempWb = new Workbook();
        Worksheet ws = tempWb.Worksheets[0];
        ws.Cells["A1"].PutValue("Header");
        ws.Cells["A1"].GetStyle().Font.IsBold = true; // Apply a style to source cell

        using (MemoryStream ms = new MemoryStream())
        {
            tempWb.Save(ms, SaveFormat.Xlsx);
            return ms.ToArray();
        }
    }
}
