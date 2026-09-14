// Title: Export rows 1‑50 and columns A‑D from an Aspose.Cells worksheet to a CSV string in C#
// AI Prompts: Generate C# code that uses Aspose.Cells TxtSaveOptions to export rows 1‑50 and columns A‑D of a worksheet into a CSV string via a MemoryStream. | Show how to configure the ExportArea property for a CSV export and read the result as a UTF‑8 string in Aspose.Cells.
// Common Searches: Aspose.Cells C# export specific range to CSV string | How to use TxtSaveOptions ExportArea for CSV in Aspose.Cells | C# get CSV output from selected worksheet area using Aspose.Cells | Export first 50 rows and columns A to D to CSV with Aspose.Cells | MemoryStream CSV conversion Aspose.Cells C# example
// Tags: Aspose.Cells CSV export of cell range | TxtSaveOptions CSV configuration | C# memory stream CSV generation | Selected worksheet area CSV output | Export worksheet subset to CSV

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example creates (or loads) a workbook, defines a TxtSaveOptions object with SaveFormat.Csv, sets ExportArea to rows 1‑50 and columns A‑D, saves the selected area to a MemoryStream, reads the stream as a UTF‑8 string, and outputs the resulting CSV content.
class ExportRowsColumnsToCsv
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // replace with new Workbook("input.xlsx") if needed

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // -----------------------------------------------------------------
        // Example data – in real scenario the worksheet would already contain data
        // -----------------------------------------------------------------
        for (int r = 0; r < 60; r++)
        {
            for (int c = 0; c < 6; c++)
            {
                sheet.Cells[r, c].PutValue($"R{r + 1}C{c + 1}");
            }
        }

        // Define the export area: rows 1‑50 (index 0‑49) and columns A‑D (index 0‑3)
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
        saveOptions.ExportArea = new CellArea
        {
            StartRow = 0,      // row 1
            EndRow = 49,       // row 50
            StartColumn = 0,   // column A
            EndColumn = 3      // column D
        };
        saveOptions.Separator = ',';   // CSV separator
        saveOptions.Encoding = Encoding.UTF8;

        // Save the selected area to a memory stream
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, saveOptions);
            ms.Position = 0;

            // Convert the stream content to a CSV string
            string csvString = new StreamReader(ms, Encoding.UTF8).ReadToEnd();

            // Output the CSV string (or use it as needed)
            Console.WriteLine(csvString);
        }
    }
}
