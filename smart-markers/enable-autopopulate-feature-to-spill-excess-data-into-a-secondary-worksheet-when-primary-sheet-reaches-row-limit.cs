// Title: Auto‑populate overflow rows to a new worksheet with Aspose.Cells for .NET
// Description: Demonstrates how to set TxtLoadOptions.ExtendToNextSheet = true so that a CSV larger than Excel's 1,048,576‑row limit automatically continues on a second worksheet, then saves the workbook as XLSX.
// Keywords: Aspose.Cells | ExtendToNextSheet | auto populate next sheet | large CSV import | Excel row limit | spill over rows | C# Aspose.Cells example
// Common Searches: Aspose.Cells auto populate next sheet C# | ExtendToNextSheet option usage | Import CSV with more than 1 million rows | Split large CSV across worksheets Aspose | How to handle Excel row limit in .NET
// Developer Intent: Enable automatic spill‑over of rows to a new worksheet when loading a CSV that exceeds the maximum rows per sheet.
// Use Cases: Loading massive data exports (>1 048 576 rows) into Excel without manual splitting. | Generating multi‑sheet reports from log files or sensor data that exceed a single sheet's capacity. | Automating data migration where source files are larger than Excel's row limit.
// AI Prompts: Show C# code that uses Aspose.Cells TxtLoadOptions to continue CSV import on a new worksheet after the row limit is reached. | Explain the effect of ExtendToNextSheet and how to verify the created worksheets. | Provide a step‑by‑step guide to import a 1.05 million‑row CSV into an XLSX file using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AutoPopulateExample
{
    // Demonstrates how to set TxtLoadOptions.ExtendToNextSheet = true so that a CSV larger than Excel's 1,048,576‑row limit automatically continues on a second worksheet, then saves the workbook as XLSX.
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Prepare a CSV data source that has more rows than a single
            //    worksheet can hold (Excel limit = 1,048,576 rows).
            //    For demonstration we generate a small CSV, but the same
            //    logic works with a huge file that exceeds the limit.
            // ------------------------------------------------------------
            var sb = new System.Text.StringBuilder();

            // Header row
            sb.AppendLine("Id,Value");

            // Generate rows – in a real scenario this would be > 1,048,576
            for (int i = 1; i <= 1_050_000; i++)
            {
                sb.AppendLine($"{i},Data_{i}");
            }

            // Convert the CSV string to a memory stream
            using var csvStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(sb.ToString()));

            // ------------------------------------------------------------
            // 2. Enable the auto‑populate (spill‑over) feature.
            //    TxtLoadOptions.ExtendToNextSheet = true tells Aspose.Cells
            //    to continue importing data into a new worksheet when the
            //    current one reaches its row limit.
            // ------------------------------------------------------------
            var loadOptions = new TxtLoadOptions
            {
                ExtendToNextSheet = true   // <-- key setting
            };

            // ------------------------------------------------------------
            // 3. Load the CSV into a Workbook using the options above.
            //    The constructor (Workbook(Stream, TxtLoadOptions)) follows
            //    the required lifecycle rule for creation/loading.
            // ------------------------------------------------------------
            var workbook = new Workbook(csvStream, loadOptions);

            // ------------------------------------------------------------
            // 4. Verify that a second worksheet was created.
            // ------------------------------------------------------------
            Console.WriteLine($"Worksheets count: {workbook.Worksheets.Count}");
            Console.WriteLine($"Rows in first sheet: {workbook.Worksheets[0].Cells.MaxDataRow + 1}");
            Console.WriteLine($"Rows in second sheet: {workbook.Worksheets[1].Cells.MaxDataRow + 1}");

            // ------------------------------------------------------------
            // 5. Save the workbook.
            //    The Save method complies with the lifecycle rule for persistence.
            // ------------------------------------------------------------
            string outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "SpillOverDemo.xlsx");
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}
