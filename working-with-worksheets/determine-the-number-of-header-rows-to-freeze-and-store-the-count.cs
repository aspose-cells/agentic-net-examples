// Title: C# – Freeze Header Rows in Excel Using Aspose.Cells TxtLoadOptions
// Description: Demonstrates how to load CSV data with Aspose.Cells, read the HeaderRowsCount from TxtLoadOptions, store the value, and freeze those rows with Worksheet.FreezePanes before saving the workbook.
// Keywords: Aspose.Cells C# freeze header rows | TxtLoadOptions HeaderRowsCount | freeze panes Aspose.Cells | load CSV Aspose.Cells .NET | Excel header freeze example
// Common Searches: how to freeze first row after importing CSV with Aspose.Cells | retrieve HeaderRowsCount from TxtLoadOptions | Aspose.Cells freeze panes based on header count | C# example freeze header rows in Excel
// Developer Intent: Capture the number of header rows defined in TxtLoadOptions and apply that count to Worksheet.FreezePanes so the headers stay visible while scrolling.
// Use Cases: Import a CSV file and automatically lock its header rows in the resulting Excel sheet. | Reuse HeaderRowsCount for styling, setting print titles, or repeating rows on each printed page. | Create a reusable routine that reads header information and applies consistent worksheet layout settings.
// AI Prompts: Write C# code that loads a CSV with Aspose.Cells, reads HeaderRowsCount from TxtLoadOptions, and freezes those rows. | Show how to adjust HeaderRowsCount dynamically based on the CSV content before calling FreezePanes.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Demonstrates how to load CSV data with Aspose.Cells, read the HeaderRowsCount from TxtLoadOptions, store the value, and freeze those rows with Worksheet.FreezePanes before saving the workbook.
class Program
{
    static void Main()
    {
        // Sample CSV data with a header row followed by data rows
        string csvData = "Header1,Header2,Header3\n" +
                         "Value1,Value2,Value3\n" +
                         "Value4,Value5,Value6";

        // Create TxtLoadOptions and set the number of header rows to be treated as headers
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            HeaderRowsCount = 1 // we have one header row
        };

        // Store the header rows count for later use
        int headerRowsToFreeze = loadOptions.HeaderRowsCount;

        // Load the CSV data into a workbook using the specified load options
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
        {
            Workbook workbook = new Workbook(stream, loadOptions);
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze the header rows so they remain visible when scrolling
            // FreezePanes(row, column, freezedRows, freezedColumns)
            worksheet.FreezePanes(headerRowsToFreeze, 0, headerRowsToFreeze, 0);

            // Save the workbook to a file
            workbook.Save("HeaderFreezeDemo.xlsx");
        }
    }
}
