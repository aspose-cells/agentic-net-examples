// Title: C# – Export rows 1‑50 and columns A‑D as a CSV string using Aspose.Cells
// Description: The sample builds a workbook, populates cells, defines a CellArea for rows 1‑50 and columns A‑D, and applies TxtSaveOptions with SaveFormat.Csv to write only that area to a MemoryStream. The stream is then read as a UTF‑8 string, yielding a CSV without creating a physical file.
// Keywords: Aspose.Cells | C# | .NET | export selected range CSV | TxtSaveOptions | ExportArea | MemoryStream | CSV string | cell area rows 1-50 | columns A-D
// Common Searches: Aspose.Cells export specific rows to CSV | C# save worksheet range as CSV string | TxtSaveOptions ExportArea example | How to convert part of a sheet to CSV in .NET | Generate CSV from selected cells Aspose.Cells
// Developer Intent: Generate a CSV representation of a defined worksheet region directly in memory.
// Use Cases: Send a CSV excerpt of the first 50 rows via an API response | Create a lightweight CSV report for email without disk I/O | Extract a subset of data from a large workbook for downstream processing
// AI Prompts: Write C# code that uses Aspose.Cells to export rows 1‑50 and columns A‑D to a CSV string using TxtSaveOptions and MemoryStream. | Explain how the ExportArea property limits CSV output to a specific cell range in Aspose.Cells for .NET. | Show how to change the output encoding when exporting a selected range to CSV with Aspose.Cells.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Saving;   // Required for TxtSaveOptions

// The sample builds a workbook, populates cells, defines a CellArea for rows 1‑50 and columns A‑D, and applies TxtSaveOptions with SaveFormat.Csv to write only that area to a MemoryStream. The stream is then read as a UTF‑8 string, yielding a CSV without creating a physical file.
class ExportRowsColumnsToCsv
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // ------------------------------------------------------------
        // Populate sample data (optional – replace with your own data)
        // ------------------------------------------------------------
        for (int row = 0; row < 60; row++)          // more than 50 rows to show the limit
        {
            for (int col = 0; col < 6; col++)       // more than 4 columns to show the limit
            {
                cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // ------------------------------------------------------------
        // Define the export area: rows 1‑50 (index 0‑49) and columns A‑D (index 0‑3)
        // ------------------------------------------------------------
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
        saveOptions.ExportArea = new CellArea
        {
            StartRow = 0,      // Row 1
            EndRow = 49,       // Row 50
            StartColumn = 0,   // Column A
            EndColumn = 3      // Column D
        };

        // ------------------------------------------------------------
        // Save the defined area to a memory stream in CSV format
        // ------------------------------------------------------------
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, saveOptions);   // Export only the specified area
            ms.Position = 0;                  // Reset stream position for reading

            // Convert the stream content to a CSV string (UTF‑8 encoding)
            string csvString = Encoding.UTF8.GetString(ms.ToArray());

            // Output the CSV string (for demonstration)
            Console.WriteLine(csvString);
        }
    }
}
