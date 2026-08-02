// Title: Aspose.Cells C# – Merge G10:H12, apply custom date format, and save as XLS via MemoryStream
// Description: Creates a new Workbook, inserts the current date into G10, sets a custom date style (dd-MMM-yyyy), merges the range G10:H12, and writes the file to a MemoryStream as an Excel 97‑2003 XLS document without requiring a physical file.
// Keywords: Aspose.Cells | C# | merge cells G10:H12 | custom date format | dd-MMM-yyyy | MemoryStream | save as XLS | Excel97-2003 | cell styling | Aspose.Cells example
// Common Searches: Aspose.Cells merge cells and set date format C# | Save Aspose.Cells workbook to MemoryStream as XLS | Apply custom date style to merged range Aspose.Cells | C# example for merging G10:H12 in Aspose.Cells | How to create Excel97-2003 file with Aspose.Cells
// Developer Intent: Generate an Excel 97‑2003 file where cells G10:H12 are merged and display the current date in a custom format, then retrieve the file from a MemoryStream.
// Use Cases: Build legacy XLS reports with a merged header that shows the generation date in a consistent format. | Return an Excel file from a web API as a byte array after merging cells and applying styling. | Create printable invoices where the date cell spans two columns and uses a locale‑specific date pattern.
// AI Prompts: Write C# code using Aspose.Cells to merge cells G10:H12, apply the date format dd-MMM-yyyy, and save the workbook to a MemoryStream as an XLS file. | Show how to set a custom date style on a merged cell range and then export the workbook to a byte array for HTTP response. | Explain step‑by‑step how to create an Excel97‑2003 workbook with merged cells and custom formatting without writing to disk first.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Creates a new Workbook, inserts the current date into G10, sets a custom date style (dd-MMM-yyyy), merges the range G10:H12, and writes the file to a MemoryStream as an Excel 97‑2003 XLS document without requiring a physical file.
    public class MergeAndFormatDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (default Xlsx format)
            using (Workbook workbook = new Workbook())
            {
                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Put a sample date value into the top-left cell of the merge range (G10)
                // Row index 9 (G10), column index 6 (G)
                worksheet.Cells[9, 6].PutValue(DateTime.Now);

                // Create a style with a custom date format and apply it to the merged cell
                Style dateStyle = worksheet.Cells[9, 6].GetStyle();
                dateStyle.Custom = "dd-MMM-yyyy"; // e.g., 27-Jul-2026
                worksheet.Cells[9, 6].SetStyle(dateStyle);

                // Merge cells G10:H12 (rows 10-12, columns G-H)
                // firstRow = 9, firstColumn = 6, totalRows = 3, totalColumns = 2
                worksheet.Cells.Merge(9, 6, 3, 2);

                // Save the workbook to a memory stream as an Excel97-2003 XLS file
                using (MemoryStream stream = workbook.SaveToStream())
                {
                    // Reset stream position if you need to read from it later
                    stream.Position = 0;

                    // Write the stream to a file for verification
                    string outputPath = "MergedDateDemo.xls";
                    using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        stream.CopyTo(file);
                    }

                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
            }
        }
    }
}
