// Title: Copy cell style between ranges using MemoryStream in Aspose.Cells for .NET (no file I/O)
// Description: Demonstrates how to create a source workbook, apply a bold Calibri style with a light‑blue background to range A1:B1, save the workbook to a MemoryStream, load a new workbook from that stream, and copy the formatting to range C1:D1—all in C# without writing any files to disk.
// Keywords: Aspose.Cells copy style | MemoryStream workbook .NET | range formatting transfer | C# Excel style copy | no disk I/O Aspose.Cells | in‑memory Excel processing | copy cell formatting between ranges
// Common Searches: Aspose.Cells copy style between ranges without file | load workbook from MemoryStream C# | copy range formatting in memory Aspose.Cells | how to transfer Excel style using streams .NET | copy cell style without saving to disk
// Developer Intent: Transfer the formatting of one cell range to another by loading workbooks from MemoryStream objects, eliminating any temporary file creation.
// Use Cases: Generate a styled template in memory and reuse its header format across multiple worksheets. | Build an in‑memory report, then apply the same style to a different sheet without intermediate files. | Process uploaded Excel files in a web API, copying styles between sheets using only streams.
// AI Prompts: Show C# code that copies a range's style to another range using Aspose.Cells with MemoryStream only. | Provide an Aspose.Cells example for transferring cell formatting between workbooks without creating temporary files. | Explain how SaveToStream and the Workbook(MemoryStream) constructor enable style copying in Aspose.Cells for .NET.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a source workbook, apply a bold Calibri style with a light‑blue background to range A1:B1, save the workbook to a MemoryStream, load a new workbook from that stream, and copy the formatting to range C1:D1—all in C# without writing any files to disk.
class Program
{
    static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Create a source workbook and apply a style to a range.
            // ------------------------------------------------------------
            Workbook srcWorkbook = new Workbook();                     // create new workbook
            Worksheet srcSheet = srcWorkbook.Worksheets[0];           // get first worksheet

            // Add some sample data
            srcSheet.Cells["A1"].PutValue("Header");
            srcSheet.Cells["A2"].PutValue("Data");

            // Define a style
            Style srcStyle = srcWorkbook.CreateStyle();
            srcStyle.Font.Name = "Calibri";
            srcStyle.Font.Size = 12;
            srcStyle.Font.IsBold = true;
            srcStyle.ForegroundColor = Color.LightBlue;
            srcStyle.Pattern = BackgroundType.Solid;

            // Apply the style to source range A1:B1
            Aspose.Cells.Range srcRange = srcSheet.Cells.CreateRange("A1:B1");
            srcRange.SetStyle(srcStyle);

            // ------------------------------------------------------------
            // 2. Save the source workbook to a memory stream (xls format).
            // ------------------------------------------------------------
            using (MemoryStream memoryStream = srcWorkbook.SaveToStream())
            {
                memoryStream.Position = 0; // reset for reading

                // ------------------------------------------------------------
                // 3. Load a new workbook from the memory stream.
                // ------------------------------------------------------------
                Workbook destWorkbook = new Workbook(memoryStream);
                Worksheet destSheet = destWorkbook.Worksheets[0];

                // ------------------------------------------------------------
                // 4. Create a destination range and copy the style from source.
                // ------------------------------------------------------------
                Aspose.Cells.Range destRange = destSheet.Cells.CreateRange("C1:D1"); // same size as source range
                destRange.CopyStyle(srcRange); // copy style

                // ------------------------------------------------------------
                // 5. (Optional) Save the result to another memory stream – no disk I/O.
                // ------------------------------------------------------------
                using (MemoryStream outStream = new MemoryStream())
                {
                    destWorkbook.Save(outStream, SaveFormat.Xlsx);
                    // outStream now contains the workbook with the copied style.
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
