// Title: Filter CSV rows by date and save the result as HTML using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells filter CSV rows by date before exporting to HTML | How to delete rows with dates earlier than a specific value in a CSV using Aspose.Cells | Convert filtered CSV data to HTML with Aspose.Cells .NET | Aspose.Cells load CSV and keep only rows after 2023-01-01 for HTML output | Remove rows with invalid dates when converting CSV to HTML in C#
// Tags: date‑based row filtering with Aspose.Cells CSV | delete rows by date criteria Aspose.Cells | export workbook to HTML using Aspose.Cells | parse and validate date cells in CSV Aspose.Cells C# | CSV to HTML conversion with row filtering Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// The example loads a CSV file into an Aspose.Cells Workbook, iterates from the last data row upward, parses the first column as a date, removes rows whose dates are earlier than January 1 2023 or cannot be parsed, and then saves the remaining rows as an HTML file.
class Program
{
    static void Main()
    {
        // Paths for input CSV and output HTML
        string csvPath = "input.csv";
        string htmlPath = "output.html";

        // Define the date filter (keep rows on or after this date)
        DateTime filterDate = new DateTime(2023, 1, 1);

        // Load the CSV file into a workbook
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
        Workbook workbook = new Workbook(csvPath, loadOptions);
        Worksheet sheet = workbook.Worksheets[0];

        // Assume the date column is the first column (A) and the first row contains headers
        int dateColumnIndex = 0; // zero‑based index for column A
        int headerRowIndex = 0;

        // Iterate rows from bottom to top and delete rows that do not meet the date criteria
        for (int row = sheet.Cells.MaxDataRow; row > headerRowIndex; row--)
        {
            string dateText = sheet.Cells[row, dateColumnIndex].StringValue;

            if (DateTime.TryParse(dateText, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime cellDate))
            {
                // Delete row if the date is earlier than the filter date
                if (cellDate < filterDate)
                {
                    sheet.Cells.DeleteRow(row);
                }
            }
            else
            {
                // If the cell cannot be parsed as a date, remove the row
                sheet.Cells.DeleteRow(row);
            }
        }

        // Save the filtered data as an HTML file
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        workbook.Save(htmlPath, htmlOptions);
    }
}
