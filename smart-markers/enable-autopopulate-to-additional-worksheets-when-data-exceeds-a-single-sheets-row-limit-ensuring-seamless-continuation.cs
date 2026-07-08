using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class AutoPopulateExample
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a CSV string that contains more rows than a single
        //    Excel worksheet can hold (1,048,576 rows). Here we generate
        //    1,050,000 rows with 5 columns.
        // ------------------------------------------------------------
        const int totalRows = 1_050_000;
        const int totalCols = 5;
        var sb = new StringBuilder();

        // Header row
        for (int c = 0; c < totalCols; c++)
        {
            sb.Append($"Col{c + 1}");
            if (c < totalCols - 1) sb.Append(',');
        }
        sb.AppendLine();

        // Data rows
        for (int r = 0; r < totalRows; r++)
        {
            for (int c = 0; c < totalCols; c++)
            {
                sb.Append($"R{r + 1}C{c + 1}");
                if (c < totalCols - 1) sb.Append(',');
            }
            sb.AppendLine();
        }

        // ------------------------------------------------------------
        // 2. Load the CSV data into a workbook.  TxtLoadOptions.ExtendToNextSheet
        //    tells Aspose.Cells to continue importing data onto a new worksheet
        //    when the current sheet reaches its row limit.
        // ------------------------------------------------------------
        using (var csvStream = new MemoryStream(Encoding.UTF8.GetBytes(sb.ToString())))
        {
            var loadOptions = new TxtLoadOptions
            {
                ExtendToNextSheet = true,   // Enable auto‑populate to additional sheets
                Separator = ','            // CSV delimiter
            };

            // Load workbook from the CSV stream with the specified options
            var workbook = new Workbook(csvStream, loadOptions);

            // ------------------------------------------------------------
            // 3. (Optional) Auto‑fit columns on each generated worksheet for
            //    better readability.
            // ------------------------------------------------------------
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.AutoFitColumns();
            }

            // ------------------------------------------------------------
            // 4. Save the workbook.  The data will be split across multiple
            //    worksheets automatically.
            // ------------------------------------------------------------
            workbook.Save("AutoPopulated.xlsx");
        }
    }
}