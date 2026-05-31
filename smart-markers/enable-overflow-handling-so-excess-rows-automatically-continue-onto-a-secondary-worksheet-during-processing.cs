using System;
using System.IO;
using Aspose.Cells;

namespace OverflowHandlingDemo
{
    class Program
    {
        static void Main()
        {
            // Prepare CSV data that exceeds the default worksheet row limit (e.g., 1,000,000 rows)
            // For demonstration we generate a simple CSV with 70,000 rows (greater than Excel 2003 limit of 65,535)
            var sb = new System.Text.StringBuilder();
            for (int i = 1; i <= 70000; i++)
            {
                sb.AppendLine($"Row{i},Value{i}");
            }
            byte[] csvBytes = System.Text.Encoding.UTF8.GetBytes(sb.ToString());

            // Configure TxtLoadOptions to extend data to the next sheet when limits are exceeded
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                ExtendToNextSheet = true   // Enable overflow handling
            };

            // Load the CSV data into a workbook using the configured options
            Workbook workbook;
            using (MemoryStream ms = new MemoryStream(csvBytes))
            {
                workbook = new Workbook(ms, loadOptions);
            }

            // At this point, excess rows are automatically placed on a new worksheet
            Console.WriteLine($"Number of worksheets created: {workbook.Worksheets.Count}");
            Console.WriteLine($"Rows in first worksheet: {workbook.Worksheets[0].Cells.MaxDataRow + 1}");
            if (workbook.Worksheets.Count > 1)
            {
                Console.WriteLine($"Rows in second worksheet: {workbook.Worksheets[1].Cells.MaxDataRow + 1}");
                Console.WriteLine($"First cell in second worksheet: {workbook.Worksheets[1].Cells[0, 0].StringValue}");
            }

            // Save the workbook to an XLSX file
            workbook.Save("OverflowHandledWorkbook.xlsx");
        }
    }
}