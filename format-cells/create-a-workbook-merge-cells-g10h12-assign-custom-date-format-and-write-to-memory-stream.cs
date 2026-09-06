// Title: Create an Excel workbook, merge cells G10:H12, apply a custom date format, and save it to a MemoryStream with Aspose.Cells for .NET
// AI Prompts: Write C# code that merges the range G10:H12 in a worksheet, inserts the current date into the merged area, formats the date as "dd-mmm-yyyy", and returns the workbook as an in‑memory XLSX stream using Aspose.Cells. | Show how to style the top‑left cell of a merged block with a custom date pattern and then export the Aspose.Cells workbook to a MemoryStream without writing to disk.
// Common Searches: Aspose.Cells C# merge a block of cells and set a custom date format | save Aspose.Cells workbook to MemoryStream after applying style to merged cells | how to apply a date style to the top‑left cell of a merged range in Aspose.Cells .NET | generate an in‑memory Excel file with merged cells containing the current date using Aspose.Cells | C# Aspose.Cells export XLSX to stream after formatting merged cell range
// Tags: merge range G10:H12 Aspose.Cells | custom date style for merged cells .NET | export workbook to MemoryStream Aspose.Cells | in‑memory XLSX creation with merged range | apply style to top‑left cell of merged block

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, merges cells G10:H12, writes the current date into the merged area, applies a "dd-mmm-yyyy" custom format, and saves the result to a MemoryStream as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells G10:H12 using a Range object
            AsposeRange mergeRange = sheet.Cells.CreateRange("G10:H12");
            mergeRange.Merge();

            // Put a date value into the merged cell (top‑left cell of the range)
            sheet.Cells["G10"].PutValue(DateTime.Now);

            // Apply a custom date format to the merged cell
            Style style = workbook.CreateStyle();
            style.Custom = "dd-mmm-yyyy";
            sheet.Cells["G10"].SetStyle(style);

            // Save the workbook to a memory stream
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.Save(stream, SaveFormat.Xlsx);
                // The MemoryStream now contains the workbook data.
                // Optionally write to a file for verification:
                // File.WriteAllBytes("MergedDate.xlsx", stream.ToArray());
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
