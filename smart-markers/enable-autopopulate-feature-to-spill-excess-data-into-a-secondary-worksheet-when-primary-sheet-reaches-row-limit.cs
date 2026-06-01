using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Prepare CSV data that exceeds Excel's maximum row limit
        //    (1,048,576 rows). Here we generate 1,050,000 rows.
        // ------------------------------------------------------------
        const int totalRows = 1_050_000;
        var sb = new StringBuilder();

        for (int i = 0; i < totalRows; i++)
        {
            // Simple two‑column CSV: Index, Value
            sb.AppendLine($"Row{i},Value{i}");
        }

        // Convert the CSV string to a UTF‑8 byte array
        byte[] csvBytes = Encoding.UTF8.GetBytes(sb.ToString());

        // ------------------------------------------------------------
        // 2. Configure TxtLoadOptions to enable auto‑populate (spill)
        //    to a new worksheet when the current sheet reaches its row limit.
        // ------------------------------------------------------------
        var loadOptions = new TxtLoadOptions
        {
            ExtendToNextSheet = true   // Enable spilling to the next sheet
        };

        // ------------------------------------------------------------
        // 3. Load the CSV data into a Workbook using the options above.
        // ------------------------------------------------------------
        using (var stream = new MemoryStream(csvBytes))
        {
            var workbook = new Workbook(stream, loadOptions);

            // Optional: adjust column widths for better readability
            workbook.Worksheets[0].AutoFitColumns();

            // ------------------------------------------------------------
            // 4. Save the resulting workbook. The excess rows will be placed
            //    in a newly created worksheet (e.g., "Sheet2").
            // ------------------------------------------------------------
            workbook.Save("SpillData.xlsx");
        }
    }
}