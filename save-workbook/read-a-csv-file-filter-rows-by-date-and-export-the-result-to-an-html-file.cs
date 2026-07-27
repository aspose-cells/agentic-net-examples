using System;
using Aspose.Cells;

namespace CsvDateFilterToHtml
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // 3. Import CSV data (rule: Cells.ImportCSV)
            //    - fileName: path to the CSV file
            //    - splitter: comma delimiter
            //    - convertNumericData: true (numeric strings become numbers)
            //    - firstRow, firstColumn: start at A1 (0,0)
            string csvPath = "input.csv"; // replace with actual CSV file path
            cells.ImportCSV(csvPath, ",", true, 0, 0);

            // 4. Apply an AutoFilter to the date column (assumed to be column A, index 0)
            //    Set the filter range to cover the whole column (adjust the row count as needed)
            //    Here we assume the data does not exceed 10000 rows.
            worksheet.AutoFilter.Range = "A1:A10000";

            // 5. Add a date filter.
            //    Example: keep only rows where the date is 2023‑05‑10.
            //    DateTimeGroupingType.Day groups by day, so only the specified day is shown.
            worksheet.AutoFilter.AddDateFilter(
                fieldIndex: 0,
                dateTimeGroupingType: DateTimeGroupingType.Day,
                year: 2023,
                month: 5,
                day: 10,
                hour: 0,
                minute: 0,
                second: 0);

            // 6. Save the filtered workbook as HTML (lifecycle: save)
            //    Use HtmlSaveOptions to export all data.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportDataOptions = HtmlExportDataOptions.All;

            workbook.Save("filtered_output.html", htmlOptions);
        }
    }
}