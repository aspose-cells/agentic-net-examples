// Title: Merge G10:H12, set custom date format, and save to MemoryStream using Aspose.Cells for .NET
// Description: Demonstrates how to create a new Workbook, insert the current date into cell G10, apply the custom format "dd-MMM-yyyy", merge the range G10:H12, and write the XLS file to a MemoryStream in C# with Aspose.Cells.
// Keywords: Aspose.Cells C# merge cells | custom date format Excel | MemoryStream workbook | save XLS to stream | in‑memory Excel file | G10:H12 merge | date header style | Aspose.Cells example
// Common Searches: Aspose.Cells merge cells and format date | save Aspose.Cells workbook to MemoryStream | C# set custom date format for merged cells | how to write Excel file to stream with Aspose.Cells | create in‑memory Excel report Aspose.Cells
// Developer Intent: Generate an in‑memory Excel workbook with a merged date header formatted as dd‑MMM‑yyyy.
// Use Cases: Produce a temporary report where the report date spans multiple columns and the file is sent as an email attachment without touching the file system. | Return an Excel spreadsheet from a web API, merging cells for a title row and applying a custom date style before streaming to the client. | Build an invoice template where the invoice date occupies a merged header area and the workbook is streamed directly to downstream services.
// AI Prompts: Write C# code with Aspose.Cells to merge cells G10:H12, apply the date format "dd-MMM-yyyy", and output the workbook as a MemoryStream. | Show an Aspose.Cells snippet that inserts the current date into a merged range, styles it, and saves the XLS file to a stream for further processing. | Explain best practices for disposing Aspose.Cells objects after saving a workbook to a MemoryStream in .NET.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a new Workbook, insert the current date into cell G10, apply the custom format "dd-MMM-yyyy", merge the range G10:H12, and write the XLS file to a MemoryStream in C# with Aspose.Cells.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Put a date value in the top-left cell of the merge area (G10)
        Cell topLeftCell = cells["G10"];
        topLeftCell.PutValue(DateTime.Now);

        // Apply a custom date format to the merged cell
        Style dateStyle = topLeftCell.GetStyle();
        dateStyle.Custom = "dd-MMM-yyyy";
        topLeftCell.SetStyle(dateStyle);

        // Merge cells G10:H12 (rows 10-12, columns G-H)
        // Zero‑based indices: row 9, column 6, spanning 3 rows and 2 columns
        cells.Merge(9, 6, 3, 2);

        // Save the workbook to a memory stream (XLS format)
        MemoryStream stream = workbook.SaveToStream();

        // Example usage of the stream (e.g., write to a file for verification)
        using (FileStream file = new FileStream("MergedDateDemo.xls", FileMode.Create, FileAccess.Write))
        {
            stream.Position = 0;
            stream.CopyTo(file);
        }

        // Clean up
        stream.Dispose();
        workbook.Dispose();
    }
}
