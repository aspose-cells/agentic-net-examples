// Title: Convert an Excel workbook to CSV with only the header row using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export only the first row of an Excel sheet to CSV | How to save Excel column headers as CSV without data using Aspose.Cells | Aspose.Cells .NET create CSV file that contains just the header row | Extract header row from .xlsx and write to CSV with Aspose.Cells in C# | Generate CSV header‑only file from Excel workbook using Aspose.Cells API
// Tags: Aspose.Cells header‑only CSV export | C# copy first worksheet row to new workbook | Save Excel workbook as CSV without data rows | Extract column headers from .xlsx using Aspose.Cells | Generate CSV file containing only headers in .NET

using System;
using Aspose.Cells;

// The example loads an input .xlsx file, reads the first worksheet, determines the number of used columns, creates a new workbook, copies only the first (header) row into it, and saves the result as a CSV file that contains just the column headers.
class WorkbookToCsvHeadersOnly
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook srcWorkbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet srcSheet = srcWorkbook.Worksheets[0];

        // Determine the number of used columns in the first row (headers)
        // MaxColumn returns the zero‑based index of the last used column
        int lastColumnIndex = srcSheet.Cells.MaxColumn;
        int columnCount = lastColumnIndex + 1; // total columns

        // Create a new workbook that will contain only the header row
        Workbook destWorkbook = new Workbook();
        Worksheet destSheet = destWorkbook.Worksheets[0];

        // Copy header values from the source sheet to the destination sheet
        for (int col = 0; col < columnCount; col++)
        {
            // Preserve the exact text of each header cell
            string headerValue = srcSheet.Cells[0, col].StringValue;
            destSheet.Cells[0, col].PutValue(headerValue);
        }

        // Save the new workbook as CSV – it will contain only the header row
        destWorkbook.Save("output.csv", SaveFormat.Csv);
    }
}
