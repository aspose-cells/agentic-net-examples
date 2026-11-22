// Title: Convert an Excel workbook to a CSV file that contains only the column headers using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export only the first row of an Excel sheet to CSV | save Excel headers as CSV without data rows using Aspose.Cells | C# code to create a CSV file with just column headings from an .xlsx workbook | how to extract header row from worksheet and write to CSV with Aspose.Cells | generate header‑only CSV from Excel using Aspose.Cells .NET library
// Tags: export header row to CSV with Aspose.Cells | save workbook as CSV containing only column headings | copy first worksheet row to new workbook C# | Aspose.Cells CSV conversion without data rows | extract Excel headers to CSV .NET

using System;
using Aspose.Cells;

// The example loads an existing Excel workbook, reads the values from the first row of the first worksheet, copies those header values into a new workbook, and saves the new workbook as a CSV file. The resulting CSV contains only the column headers and no data rows.
class Program
{
    static void Main()
    {
        // Path to the source Excel workbook
        string inputPath = @"C:\Input\SourceWorkbook.xlsx";

        // Path where the CSV with only headers will be saved
        string outputPath = @"C:\Output\HeadersOnly.csv";

        // Load the source workbook (using the provided load rule)
        Workbook sourceWorkbook = new Workbook(inputPath);

        // Get the first worksheet (you can change the index if needed)
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Determine the last column that contains a header (non‑empty cell in the first row)
        int lastColumn = sourceSheet.Cells.MaxColumn;

        // Create a new workbook that will hold only the header row (using the provided create rule)
        Workbook headerOnlyWorkbook = new Workbook();

        // Get the default worksheet in the new workbook
        Worksheet headerSheet = headerOnlyWorkbook.Worksheets[0];

        // Copy each header cell from the source sheet to the new sheet
        for (int col = 0; col <= lastColumn; col++)
        {
            // Read the header value from the source (row 0)
            object headerValue = sourceSheet.Cells[0, col].Value;

            // Write the header value to the new workbook (row 0)
            headerSheet.Cells[0, col].PutValue(headerValue);
        }

        // Save the new workbook as CSV (using the provided save rule)
        // The SaveFormat.Csv ensures the file is written in CSV format.
        headerOnlyWorkbook.Save(outputPath, SaveFormat.Csv);
    }
}
