// Title: Determine Header Column Count and Freeze Columns with Aspose.Cells for .NET
// Description: Loads CSV data into an Aspose.Cells workbook using TxtLoadOptions, reads the HeaderColumnsCount value, freezes the same number of columns with FreezePanes, and saves the result as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | FreezePanes | HeaderColumnsCount | TxtLoadOptions | CSV import | freeze columns | worksheet freeze panes | Excel automation example
// Common Searches: Aspose.Cells freeze first columns | HeaderColumnsCount usage in C# | How to freeze columns based on header count | Load CSV with TxtLoadOptions Aspose.Cells | Freeze panes programmatically .NET
// Developer Intent: Read the header column count from load options and apply FreezePanes to lock those columns in the worksheet.
// Use Cases: Import a CSV where the first N columns are headers and keep them visible while scrolling. | Dynamically determine the number of header columns from TxtLoadOptions and apply a column freeze. | Generate Excel reports that require frozen header columns for better readability.
// AI Prompts: Write C# code that loads a CSV with Aspose.Cells, sets HeaderColumnsCount, retrieves the count, and freezes those columns. | Explain the relationship between TxtLoadOptions.HeaderColumnsCount and FreezePanes in Aspose.Cells with a short example. | Show how to freeze both header rows and columns using Aspose.Cells based on configurable counts.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHeaderFreezeDemo
{
    // Loads CSV data into an Aspose.Cells workbook using TxtLoadOptions, reads the HeaderColumnsCount value, freezes the same number of columns with FreezePanes, and saves the result as an Excel file.
    class Program
    {
        static void Main()
        {
            // Sample CSV data with header columns
            string csvData = "Header1,Header2,Header3,Data1,Data2,Data3\n" +
                             "H1,H2,H3,D1,D2,D3\n" +
                             "A,B,C,D,E,F\n" +
                             "G,H,I,J,K,L";

            // Define how many columns are considered headers
            int headerColumnsToTreatAsHeader = 3; // example value

            // Create TxtLoadOptions and set HeaderColumnsCount
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                HeaderColumnsCount = headerColumnsToTreatAsHeader,
                Encoding = Encoding.UTF8
            };

            // Load the CSV data into a workbook using the options
            using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
            {
                Workbook workbook = new Workbook(stream, loadOptions);
                Worksheet worksheet = workbook.Worksheets[0];

                // Determine the number of header columns (store the count)
                int headerCount = loadOptions.HeaderColumnsCount;

                // Freeze the header columns in the worksheet
                // Freeze at column index = headerCount (0‑based), row index = 0
                // Freeze only the left pane columns (no rows frozen)
                worksheet.FreezePanes(0, headerCount, 0, headerCount);

                // (Optional) Demonstrate that the freeze was applied
                Console.WriteLine($"Header columns frozen: {headerCount}");

                // Save the workbook to a file
                workbook.Save("HeaderFreezeDemo.xlsx");
            }
        }
    }
}
