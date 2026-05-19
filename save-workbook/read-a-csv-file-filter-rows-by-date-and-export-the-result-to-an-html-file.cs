using System;
using System.IO;
using Aspose.Cells;

class CsvDateFilterToHtml
{
    static void Main()
    {
        try
        {
            // Paths – adjust as needed
            string csvPath = "input.csv";
            string htmlPath = "filtered_output.html";

            // Verify that the CSV file exists to avoid FileNotFoundException
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"Error: CSV file not found at '{csvPath}'.");
                return;
            }

            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Import CSV data (comma‑separated, convert numeric data)
            cells.ImportCSV(csvPath, ",", true, 0, 0);

            // 3. Determine the used range (assumes dates are in column A)
            int lastRow = cells.MaxDataRow;                     // zero‑based index of last used row
            string range = $"A1:A{lastRow + 1}";                // Excel range string (1‑based rows)

            // 4. Apply an AutoFilter to the date column (field index 0)
            sheet.AutoFilter.Range = range;

            // 5. Filter rows where the date is greater than a specific date (e.g., 2023‑01‑01)
            DateTime filterDate = new DateTime(2023, 1, 1);
            sheet.AutoFilter.Custom(0, FilterOperatorType.GreaterThan, filterDate);
            sheet.AutoFilter.Refresh(); // Apply the filter

            // 6. Save the filtered worksheet as HTML, exporting all data
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportDataOptions = HtmlExportDataOptions.All // export full data
            };
            workbook.Save(htmlPath, htmlOptions);

            Console.WriteLine($"Filtered HTML saved to '{htmlPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}