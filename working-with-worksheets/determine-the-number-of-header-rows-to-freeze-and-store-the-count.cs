// Title: Freeze Header Rows in a CSV with Aspose.Cells for .NET and Capture the Count
// Description: Demonstrates how to load CSV data into an Aspose.Cells Workbook, set TxtLoadOptions.HeaderRowsCount, freeze a configurable number of header rows using FreezePanes, save the file, and output the frozen row count.
// Keywords: Aspose.Cells C# freeze rows | TxtLoadOptions HeaderRowsCount | FreezePanes CSV .NET | store frozen header count | Excel header freeze Aspose | C# load CSV memory stream | global .NET developers
// Common Searches: how to freeze the first N rows in Excel with Aspose.Cells | set header row count when importing CSV using Aspose.Cells | retrieve number of frozen rows after FreezePanes | C# example for freezing panes based on variable
// Developer Intent: Apply a variable number of frozen header rows to a worksheet and keep the count accessible in code.
// Use Cases: Import a CSV, define how many rows are headers, and keep them visible while scrolling. | Allow users to specify the header depth at runtime and automatically freeze those rows. | Log or display the exact number of rows that were frozen after workbook generation.
// AI Prompts: Write C# code that reads CSV text, uses TxtLoadOptions.HeaderRowsCount, freezes the specified header rows with Aspose.Cells, and returns the frozen row count. | Show how to dynamically set FreezePanes parameters based on a user‑provided header row number. | Explain how to obtain and print the count of frozen rows from a worksheet after calling FreezePanes.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHeaderFreezeDemo
{
    // Demonstrates how to load CSV data into an Aspose.Cells Workbook, set TxtLoadOptions.HeaderRowsCount, freeze a configurable number of header rows using FreezePanes, save the file, and output the frozen row count.
    public class Program
    {
        public static void Main()
        {
            // Sample CSV data with a header row and some data rows
            string csvData = "Header1,Header2,Header3\n" +
                             "Data1,Data2,Data3\n" +
                             "Data4,Data5,Data6\n" +
                             "Data7,Data8,Data9";

            // Define how many header rows should be treated as headers
            int headerRowsToFreeze = 1;

            // Configure TxtLoadOptions to recognize the header rows
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                HeaderRowsCount = headerRowsToFreeze
            };

            // Load the CSV data into a workbook using the specified options
            using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
            {
                Workbook workbook = new Workbook(stream, loadOptions);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Freeze the header rows so they stay visible while scrolling
                // FreezePanes(rowIndex, columnIndex, frozenRows, frozenColumns)
                // Row index is zero‑based; to freeze the first N rows, use row = N
                worksheet.FreezePanes(headerRowsToFreeze, 0, headerRowsToFreeze, 0);

                // Save the workbook (adjust path/format as needed)
                workbook.Save("HeaderRowsFrozen.xlsx");
            }

            // The variable headerRowsToFreeze now holds the count of frozen header rows
            Console.WriteLine($"Number of header rows frozen: {headerRowsToFreeze}");
        }
    }
}
