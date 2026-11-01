// Title: Delete multiple non‑contiguous row ranges in a worksheet with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# delete rows from several non‑adjacent ranges | How to remove separate row blocks in an Excel file using Aspose.Cells .NET | C# example for deleting rows A1:B2 and D4:E5 with Aspose.Cells | Delete non‑contiguous rows in an Excel workbook via Aspose.Cells API | Aspose.Cells DeleteRows for multiple row intervals without shifting data
// Tags: Aspose.Cells DeleteRows non‑contiguous ranges | C# delete multiple row blocks Excel | Aspose.Cells delete rows by index range | Excel row removal descending order Aspose.Cells | Aspose.Cells .NET remove separate row intervals

using System;
using System.IO;
using Aspose.Cells;

namespace DeleteNonContiguousRangesApp
{
    // The example creates a workbook, writes sample data to cells A1:B2 and D4:E5, then deletes the two non‑contiguous row blocks by calling DeleteRows first for rows 3‑4 and then for rows 0‑1 to keep indices correct, and finally saves the file as DeletedNonContiguousRanges.xlsx.
    class DeleteNonContiguousRanges
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (optional, just for demonstration)
                sheet.Cells["A1"].PutValue("R1C1");
                sheet.Cells["B1"].PutValue("R1C2");
                sheet.Cells["A2"].PutValue("R2C1");
                sheet.Cells["B2"].PutValue("R2C2");
                sheet.Cells["D4"].PutValue("R4C4");
                sheet.Cells["E4"].PutValue("R4C5");
                sheet.Cells["D5"].PutValue("R5C4");
                sheet.Cells["E5"].PutValue("R5C5");

                // Delete rows belonging to the two non‑contiguous ranges:
                // Range A1:B2 corresponds to rows 0‑1
                // Range D4:E5 corresponds to rows 3‑4 (0‑based)
                // Delete in descending order to keep indices correct
                sheet.Cells.DeleteRows(3, 2); // Deletes rows 3 and 4 (original D4:E5)
                sheet.Cells.DeleteRows(0, 2); // Deletes rows 0 and 1 (original A1:B2)

                // Define output file path
                string outputPath = "DeletedNonContiguousRanges.xlsx";

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
