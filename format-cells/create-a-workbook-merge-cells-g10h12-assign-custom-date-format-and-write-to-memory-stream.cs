// Title: Merge cells G10:H12, apply a custom date format, and save workbook to a MemoryStream using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge a range and set a custom date format | write an Excel workbook to a MemoryStream with Aspose.Cells | set style for merged cells in Aspose.Cells .NET | export workbook with merged date cell to XLSX stream using Aspose.Cells | apply dd-MMM-yyyy format to merged cell in Aspose.Cells
// Tags: merge cell range G10:H12 Aspose.Cells | custom date format style Aspose.Cells | save workbook as XLSX to MemoryStream | apply style to merged cell C# | insert current date into merged cell Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, merges cells G10:H12, applies the custom date format dd-MMM-yyyy to the merged cell, writes the current date, and saves the workbook as an XLSX file into a MemoryStream.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells G10:H12 (zero‑based indices: row 9, column 6, 3 rows, 2 columns)
        sheet.Cells.Merge(9, 6, 3, 2);

        // Set a custom date format for the merged cell (G10)
        Cell mergedCell = sheet.Cells["G10"];
        Style style = mergedCell.GetStyle();
        style.Custom = "dd-MMM-yyyy";   // Example: 25-Dec-2023
        mergedCell.SetStyle(style);

        // Optionally write a date value to the merged cell
        mergedCell.PutValue(DateTime.Now);

        // Save the workbook to a memory stream in XLSX format
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, SaveFormat.Xlsx);
            // The memory stream now contains the workbook data.
            // Reset position if you need to read from it later.
            ms.Position = 0;

            // Example: write the stream to a file (remove if not needed)
            // File.WriteAllBytes("MergedDateWorkbook.xlsx", ms.ToArray());
        }
    }
}
