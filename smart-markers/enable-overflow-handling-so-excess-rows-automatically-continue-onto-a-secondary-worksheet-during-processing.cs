// Title: Aspose.Cells for .NET: Auto‑split CSV rows beyond 65,535 into additional worksheets using TxtLoadOptions.ExtendToNextSheet
// Description: Demonstrates how to load a CSV with 70,000 rows in C# using Aspose.Cells, enable TxtLoadOptions.ExtendToNextSheet, and let the library automatically create a second worksheet for rows that exceed the legacy Excel limit. The workbook is saved as OverflowResult.xlsx.
// Keywords: Aspose.Cells | TxtLoadOptions | ExtendToNextSheet | C# | .NET | large CSV import | Excel row limit 65535 | multiple worksheets | overflow handling | GitHub example | US developers | Europe developers
// Common Searches: Aspose.Cells ExtendToNextSheet example C# | How to split CSV rows into new worksheets with Aspose.Cells | Load CSV larger than 65k rows into Excel using .NET | Overflow handling for large CSV files Aspose.Cells | GitHub Aspose.Cells CSV overflow sample
// Developer Intent: Configure TxtLoadOptions so that rows exceeding Excel's maximum are automatically placed on a new worksheet during CSV import.
// Use Cases: Import massive CSV logs (>65,535 rows) and let Aspose.Cells generate extra sheets without manual splitting. | Create Excel reports from big data sets where each sheet must stay within Excel's row capacity. | Automate conversion of large delimited files to a multi‑sheet workbook in a .NET service.
// AI Prompts: Show a C# code snippet that uses TxtLoadOptions.ExtendToNextSheet to divide a 70k‑row CSV into multiple worksheets with Aspose.Cells. | Explain the behavior, limits, and configuration options of ExtendToNextSheet for different delimiters and encodings. | Provide step‑by‑step instructions to verify row counts per worksheet after overflow handling and save the workbook.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to load a CSV with 70,000 rows in C# using Aspose.Cells, enable TxtLoadOptions.ExtendToNextSheet, and let the library automatically create a second worksheet for rows that exceed the legacy Excel limit. The workbook is saved as OverflowResult.xlsx.
class OverflowDemo
{
    static void Main()
    {
        // Generate CSV data with more rows than the legacy Excel limit (65535 rows)
        int totalRows = 70000;
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < totalRows; i++)
        {
            sb.AppendLine($"Row{i + 1},Value{i + 1}");
        }
        byte[] csvBytes = System.Text.Encoding.UTF8.GetBytes(sb.ToString());

        // Enable overflow handling: excess rows will be placed on a new worksheet
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            ExtendToNextSheet = true
        };

        // Load the CSV data into a workbook using the specified options
        using (MemoryStream ms = new MemoryStream(csvBytes))
        {
            Workbook workbook = new Workbook(ms, loadOptions);

            // Output information about the created worksheets
            Console.WriteLine($"Worksheets count: {workbook.Worksheets.Count}");
            Console.WriteLine($"Rows in first worksheet: {workbook.Worksheets[0].Cells.MaxDataRow + 1}");
            if (workbook.Worksheets.Count > 1)
            {
                Console.WriteLine($"Rows in second worksheet: {workbook.Worksheets[1].Cells.MaxDataRow + 1}");
            }

            // Save the result
            workbook.Save("OverflowResult.xlsx");
        }
    }
}
