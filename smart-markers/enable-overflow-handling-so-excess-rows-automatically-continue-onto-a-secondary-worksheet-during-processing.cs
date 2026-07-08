using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsOverflowDemo
{
    class Program
    {
        static void Main()
        {
            // Generate CSV data with more rows than a single worksheet can hold (e.g., 70,000 rows)
            // Excel worksheet limit is 1,048,576 rows, but we use a smaller limit for demonstration.
            int totalRows = 70000;
            var csvBuilder = new System.Text.StringBuilder();
            csvBuilder.AppendLine("Index,Value"); // header
            for (int i = 1; i <= totalRows; i++)
            {
                csvBuilder.AppendLine($"{i},Data_{i}");
            }

            // Convert CSV string to a memory stream
            using (var csvStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(csvBuilder.ToString())))
            {
                // Enable overflow handling so excess rows continue on the next worksheet
                TxtLoadOptions loadOptions = new TxtLoadOptions
                {
                    ExtendToNextSheet = true
                };

                // Load the CSV data into a workbook using the specified options
                Workbook workbook = new Workbook(csvStream, loadOptions);

                // Optional: display information about the result
                Console.WriteLine($"Worksheets created: {workbook.Worksheets.Count}");
                Console.WriteLine($"Rows in first worksheet: {workbook.Worksheets[0].Cells.MaxDataRow + 1}");
                if (workbook.Worksheets.Count > 1)
                {
                    Console.WriteLine($"Rows in second worksheet: {workbook.Worksheets[1].Cells.MaxDataRow + 1}");
                }

                // Save the workbook to an Excel file
                workbook.Save("OverflowResult.xlsx");
            }
        }
    }
}