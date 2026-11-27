// Title: Export rows 1‑50 and columns A‑D from an Excel worksheet to a CSV string using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export range A1:D50 to CSV string | Convert Excel cells A1 to D50 into CSV without writing a file in .NET | How to use ExportDataTable to generate CSV from specific rows and columns | Generate in‑memory CSV from worksheet rows 1‑50 columns A‑D using Aspose.Cells | C# example for exporting a selected range to CSV with proper field escaping
// Tags: Aspose.Cells export selected range to CSV | ExportDataTable to CSV conversion .NET | in‑memory CSV generation from Excel data | CSV field escaping for Excel export | range A1:D50 extraction Aspose.Cells

using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Cells;

namespace ExportRangeToCsvString
{
    // The sample loads (or creates) an Excel workbook, selects the A1:D50 range on the first worksheet, exports the range to a DataTable via ExportDataTable, converts the DataTable to a properly escaped CSV string in memory, prints the first few lines, and disposes the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to an existing Excel file (optional). If not provided, a sample workbook will be created.
            string excelPath = args.Length > 0 ? args[0] : null;

            Workbook workbook = null;

            try
            {
                // Load workbook from file if the path is valid; otherwise create a sample workbook.
                if (!string.IsNullOrEmpty(excelPath) && File.Exists(excelPath))
                {
                    workbook = new Workbook(excelPath);
                    Console.WriteLine($"Loaded workbook from \"{excelPath}\".");
                }
                else
                {
                    workbook = new Workbook();
                    PopulateSampleData(workbook);
                    Console.WriteLine("Created a sample workbook with dummy data.");
                }

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the range: rows 1‑50 and columns A‑D => "A1:D50".
                string rangeAddress = "A1:D50";

                // Resolve ambiguity between Aspose.Cells.Range and System.Range.
                Aspose.Cells.Range range = worksheet.Cells.CreateRange(rangeAddress);

                // Export the range to a DataTable.
                DataTable dataTable = range.ExportDataTable();

                // Convert DataTable to CSV string.
                string csvString = ConvertDataTableToCsv(dataTable);

                // Output the CSV string (first 10 lines).
                Console.WriteLine("=== CSV Output (first 10 lines) ===");
                using (StringReader sr = new StringReader(csvString))
                {
                    for (int i = 0; i < 10; i++)
                    {
                        string line = sr.ReadLine();
                        if (line == null) break;
                        Console.WriteLine(line);
                    }

                    if (sr.Peek() != -1)
                    {
                        Console.WriteLine("... (truncated)");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                workbook?.Dispose();
            }
        }

        private static void PopulateSampleData(Workbook workbook)
        {
            Worksheet sheet = workbook.Worksheets[0];

            // Header
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Date");
            sheet.Cells["D1"].PutValue("Amount");
            sheet.Cells["E1"].PutValue("Notes");

            // Sample rows
            for (int i = 2; i <= 60; i++)
            {
                sheet.Cells[$"A{i}"].PutValue(i - 1);
                sheet.Cells[$"B{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"C{i}"].PutValue(DateTime.Today.AddDays(i - 2).ToShortDateString());
                sheet.Cells[$"D{i}"].PutValue((i - 1) * 10.5);
                sheet.Cells[$"E{i}"].PutValue(i % 5 == 0 ? $"Special, note \"{i}\"" : $"Note {i}");
            }
        }

        private static string ConvertDataTableToCsv(DataTable table)
        {
            StringBuilder sb = new StringBuilder();

            // Uncomment the following lines if column headers are required.
            //var headerFields = table.Columns.Cast<DataColumn>()
            //    .Select(col => EscapeCsvField(col.ColumnName));
            //sb.AppendLine(string.Join(",", headerFields));

            foreach (DataRow row in table.Rows)
            {
                var fields = row.ItemArray.Select(field => EscapeCsvField(field?.ToString() ?? string.Empty));
                sb.AppendLine(string.Join(",", fields));
            }

            return sb.ToString();
        }

        /// <summary>
        /// Escapes a single CSV field according to RFC 4180.
        /// </summary>
        private static string EscapeCsvField(string field)
        {
            bool mustQuote = field.Contains(',') || field.Contains('"') || field.Contains('\r') || field.Contains('\n');
            if (mustQuote)
            {
                // Escape double quotes by doubling them.
                string escaped = field.Replace("\"", "\"\"");
                return $"\"{escaped}\"";
            }
            else
            {
                return field;
            }
        }
    }
}
