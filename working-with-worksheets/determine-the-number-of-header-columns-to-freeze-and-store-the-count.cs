// Title: Freeze Header Columns in a Worksheet with Aspose.Cells for .NET and Capture the Count
// Description: Demonstrates loading CSV data into an Aspose.Cells Workbook using TxtLoadOptions, setting HeaderColumnsCount, freezing the specified number of columns with FreezePanes, storing the frozen column count in a variable, and saving the result as an XLSX file.
// Keywords: Aspose.Cells freeze columns | HeaderColumnsCount .NET | TxtLoadOptions CSV import | FreezePanes example C# | store frozen column count | Aspose.Cells worksheet freezing | C# load CSV Aspose.Cells | Excel column freeze programmatically
// Common Searches: how to freeze first N columns using Aspose.Cells | HeaderColumnsCount usage in Aspose.Cells | freeze panes based on header count C# | store number of frozen columns Aspose.Cells | load CSV and freeze header columns .NET
// Developer Intent: Freeze a defined number of header columns and retain the count for later processing.
// Use Cases: Import a CSV file, treat the initial columns as headers, and keep them visible while scrolling horizontally. | Apply consistent column‑freeze settings across multiple worksheets after merging data from different CSV sources. | Use the stored header count to adjust UI layouts, generate reports, or drive further calculations that depend on the frozen columns.
// AI Prompts: Generate C# code that reads a CSV with Aspose.Cells, freezes the first N columns based on HeaderColumnsCount, and returns the frozen column count. | Show how to refactor the sample so the header column count is passed as a method argument and applied to every worksheet in a workbook. | Provide an example that writes the frozen header column count to a JSON configuration file after applying FreezePanes.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Demonstrates loading CSV data into an Aspose.Cells Workbook using TxtLoadOptions, setting HeaderColumnsCount, freezing the specified number of columns with FreezePanes, storing the frozen column count in a variable, and saving the result as an XLSX file.
class Program
{
    static void Main()
    {
        // Sample CSV data
        string csvData = "Header1,Header2,Header3,Data1,Data2\n" +
                         "H1,H2,H3,D1,D2\n" +
                         "A,B,C,D,E\n" +
                         "G,H,I,J,K";

        // Configure TxtLoadOptions with the desired number of header columns
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            HeaderColumnsCount = 3,          // Number of columns to treat as headers
            Encoding = Encoding.UTF8
        };

        // Load the workbook from the CSV stream using the specified options
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
        {
            Workbook workbook = new Workbook(stream, loadOptions);
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the number of header columns to freeze
            int headerColumnsToFreeze = loadOptions.HeaderColumnsCount;

            // Freeze the header columns (no rows are frozen)
            // row index = 0, column index = headerColumnsToFreeze,
            // freezedRows = 0, freezedColumns = headerColumnsToFreeze
            worksheet.FreezePanes(0, headerColumnsToFreeze, 0, headerColumnsToFreeze);

            // Store the count for later use (example variable)
            int frozenHeaderColumnCount = headerColumnsToFreeze;

            // Save the workbook
            workbook.Save("HeaderFreezeDemo.xlsx");
        }
    }
}
