// Title: Convert an Excel table to a regular range, preserve the first five rows' formatting, and export to ODS with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an XLSX workbook, removes the first ListObject while preserving its data, clears styles for rows after the fifth, and saves the file as ODS using Aspose.Cells. | Show how to keep the original formatting of the top five rows when converting an Excel table to a normal range and then export to ODS with Aspose.Cells. | Demonstrate resetting cell styles beyond a specific row index after a table‑to‑range conversion in Aspose.Cells for .NET.
// Common Searches: asp.net convert excel table to range keep formatting first five rows and export as ods | c# aspose.cells delete ListObject but retain data then save workbook in ods format | how to clear cell formatting after row 5 after converting a table to a range using Aspose.Cells | save workbook as ODS after converting Excel table to normal range with Aspose.Cells .NET
// Tags: Aspose.Cells table conversion preserving top rows formatting | Aspose.Cells ODS export after range transformation | Aspose.Cells reset cell styles beyond specific row | Aspose.Cells preserve worksheet data when removing a table | Aspose.Cells transform table into regular range

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The example loads an XLSX file, extracts the first ListObject, converts it to a standard range, clears formatting for all rows beyond the fifth while keeping the original styles for the top five rows, and finally saves the modified workbook as an ODS file using Aspose.Cells for .NET.
class TableToRangeExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.ods";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one table (ListObject) on the sheet
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No table found on the worksheet.");
                return;
            }

            // Get the first table
            ListObject table = sheet.ListObjects[0];

            // Store the data range before removing the table
            AsposeRange dataRange = table.DataRange;

            // Remove the table but keep the data (converts it to a normal range)
            sheet.ListObjects.RemoveAt(0);

            // Determine range boundaries
            int startRow = dataRange.FirstRow;
            int endRow = dataRange.FirstRow + dataRange.RowCount - 1;
            int startCol = dataRange.FirstColumn;
            int endCol = dataRange.FirstColumn + dataRange.ColumnCount - 1;

            // Create a default (empty) style once
            Style defaultStyle = workbook.CreateStyle();

            // Clear formatting for rows beyond the first five
            for (int row = startRow + 5; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    sheet.Cells[row, col].SetStyle(defaultStyle);
                }
            }

            // Save the modified workbook as ODS
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
