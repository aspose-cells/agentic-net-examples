// Title: Auto‑populate additional worksheets when CSV/TSV rows exceed Excel limit using Aspose.Cells for .NET
// Description: This example shows how to split a massive data set across multiple worksheets. It first demonstrates manual paging of 1,050,000 rows, then uses TxtLoadOptions with ExtendToNextSheet=true to let Aspose.Cells automatically create new sheets for overflow rows when loading a CSV/TSV file, followed by AutoFitColumns and saving the workbooks.
// Keywords: Aspose.Cells .NET | ExtendToNextSheet | CSV overflow to new sheet | TSV large dataset | Excel row limit 1,048,576 | auto split worksheets | auto‑fit columns | large data export | multiple worksheets handling | Aspose.Cells TxtLoadOptions
// Common Searches: Aspose.Cells create new worksheet when CSV exceeds row limit | ExtendToNextSheet option example | split large CSV into multiple Excel sheets .NET | auto‑populate worksheets after loading TSV with Aspose.Cells | handle Excel 1,048,576 row limit in C#
// Developer Intent: Automatically continue data on new worksheets once a sheet reaches Excel's maximum row count.
// Use Cases: Import a multi‑gigabyte CSV file and let Aspose.Cells generate additional worksheets without manual code. | Generate reports that may exceed one sheet, preserving header rows on each new worksheet. | Programmatically split any large data collection across worksheets while keeping column widths optimal.
// AI Prompts: Provide C# code that loads a CSV with Aspose.Cells and uses ExtendToNextSheet to create extra worksheets automatically. | Explain how TxtLoadOptions.ExtendToNextSheet works and its limitations with TSV files. | Show how to manually split a dataset across worksheets and then compare it with the automatic overflow feature.

using System;
using System.IO;
using Aspose.Cells;

// This example shows how to split a massive data set across multiple worksheets. It first demonstrates manual paging of 1,050,000 rows, then uses TxtLoadOptions with ExtendToNextSheet=true to let Aspose.Cells automatically create new sheets for overflow rows when loading a CSV/TSV file, followed by AutoFitColumns and saving the workbooks.
class Program
{
    static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // Part 1: Manually split a large data set across multiple sheets
            // ------------------------------------------------------------

            const int MaxRowsPerSheet = 1_048_576; // Excel row limit per worksheet

            // Create a new workbook and set up the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data_1";

            // Write header
            sheet.Cells[0, 0].PutValue("ID");
            sheet.Cells[0, 1].PutValue("Description");

            int currentRow = 1;   // start after header
            int sheetIndex = 0;

            // Simulated data source exceeding one sheet
            int totalRows = 1_050_000;

            for (int i = 1; i <= totalRows; i++)
            {
                // Create a new sheet when the current one is full
                if (currentRow >= MaxRowsPerSheet)
                {
                    sheetIndex = workbook.Worksheets.Add();
                    sheet = workbook.Worksheets[sheetIndex];
                    sheet.Name = $"Data_{sheetIndex + 1}";

                    // Write header on the new sheet
                    sheet.Cells[0, 0].PutValue("ID");
                    sheet.Cells[0, 1].PutValue("Description");

                    currentRow = 1;
                }

                // Populate data row
                sheet.Cells[currentRow, 0].PutValue(i);
                sheet.Cells[currentRow, 1].PutValue($"Item {i}");
                currentRow++;
            }

            // Auto‑fit columns for readability
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.AutoFitColumns();
            }

            workbook.Save("LargeDataSplit.xlsx");

            // ------------------------------------------------------------
            // Part 2: Load a CSV/TSV file and let Aspose.Cells auto‑populate
            //         additional worksheets when the row limit is exceeded
            // ------------------------------------------------------------

            string csvFilePath = "largeData.csv";

            // Ensure the CSV file exists; create a simple one if missing
            if (!File.Exists(csvFilePath))
            {
                using (StreamWriter sw = new StreamWriter(csvFilePath))
                {
                    sw.WriteLine("ID,Description");
                    for (int i = 1; i <= 10; i++)
                    {
                        sw.WriteLine($"{i},Item {i}");
                    }
                }
            }

            // Configure load options to extend overflow rows to new sheets
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                ExtendToNextSheet = true
            };

            // Load the CSV file
            Workbook csvWorkbook = new Workbook(csvFilePath, loadOptions);

            // Auto‑fit columns in the loaded workbook
            foreach (Worksheet ws in csvWorkbook.Worksheets)
            {
                ws.AutoFitColumns();
            }

            csvWorkbook.Save("CsvDataWithOverflow.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
