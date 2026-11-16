// Title: Freeze the first row and column of a worksheet and save the workbook to a MemoryStream with Aspose.Cells for .NET (C#)
// Common Searches: C# code to freeze top row and left column using Aspose.Cells | How to export an Aspose.Cells workbook to a MemoryStream as XLSX | Aspose.Cells FreezePanes example for .NET developers | Saving Excel file to stream instead of disk with Aspose.Cells C# | Write Excel workbook to byte array using Aspose.Cells memory stream
// Tags: Aspose.Cells freeze panes C# | write workbook to MemoryStream Aspose.Cells | export XLSX to stream .NET | freeze first row and column Aspose.Cells | Aspose.Cells in-memory Excel generation

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a new Workbook, accesses its first Worksheet, freezes the first row and column with FreezePanes(1,1,1,1), and then saves the workbook as an XLSX file into a MemoryStream for in‑memory processing.
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

                // Freeze the first row and the first column.
                // FreezePanes(row, column, totalRows, totalColumns) uses zero‑based indices.
                // row = 1, column = 1 -> split after the first row/column.
                // totalRows = 1, totalColumns = 1 -> freeze that first row and column.
                sheet.FreezePanes(1, 1, 1, 1);

                // Save the workbook to a memory stream
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                    // The MemoryStream now contains the XLSX file data.
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
