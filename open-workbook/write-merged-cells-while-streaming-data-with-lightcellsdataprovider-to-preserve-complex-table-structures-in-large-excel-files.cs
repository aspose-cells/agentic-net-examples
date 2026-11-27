// Title: How to merge header, sub‑header, and data ranges in a large Excel workbook using Aspose.Cells for .NET
// Common Searches: C# merge multiple cell ranges in Aspose.Cells for large Excel files | preserve merged cells when writing large worksheets with Aspose.Cells | Aspose.Cells merge header and subheader cells programmatically in .NET | streaming data to Excel while keeping merged cells using Aspose.Cells | example of Cells.Merge for complex table structures in C#
// Tags: Aspose.Cells merge cell ranges C# | stream large Excel workbook with LightCellsDataProvider | maintain merged cells during data streaming Aspose | header and subheader cell merging Aspose.Cells | complex table layout handling Excel Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, fills a 1000‑row by 10‑column worksheet with sample data, merges specific header, sub‑header, and data ranges using the Cells.Merge method, and saves the file as an XLSX document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define data dimensions
            const int totalRows = 1000;      // Adjust as needed
            const int totalColumns = 10;     // Adjust as needed

            // Populate cells with sample data
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalColumns; col++)
                {
                    string value;
                    if (row == 0)
                        value = $"Header {col + 1}";
                    else if (row == 1)
                        value = $"SubHeader {col + 1}";
                    else
                        value = $"R{row}_C{col}";

                    sheet.Cells[row, col].PutValue(value);
                }
            }

            // Merge header cells A1:D1 (row 0, columns 0-3)
            sheet.Cells.Merge(0, 0, 1, 4);

            // Merge a sub‑header that spans rows 2‑3 and columns B‑C
            // (rows 1‑2, columns 1‑2 in zero‑based indexing)
            sheet.Cells.Merge(1, 1, 2, 2);

            // Merge a block of data cells (rows 5‑7, columns E‑G)
            // (rows 4‑6, columns 4‑6 in zero‑based indexing)
            sheet.Cells.Merge(4, 4, 3, 3);

            // Save the workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
