// Title: C# – Filter CSV rows by a specific date with Aspose.Cells AutoFilter and export to HTML
// Description: This example shows how to load a CSV file into an Aspose.Cells workbook, apply an AutoFilter that keeps only rows matching a given DateTime (using Day grouping), and save the filtered result as an HTML page with all data exported via HtmlSaveOptions.
// Keywords: Aspose.Cells CSV import | AutoFilter date filter .NET | DateTimeGroupingType.Day | Export workbook to HTML | C# filter CSV by date | HtmlSaveOptions ExportDataOptions.All | Aspose.Cells tutorial
// Common Searches: Aspose.Cells filter CSV by exact date C# | How to export filtered CSV data to HTML using Aspose.Cells | AutoFilter date grouping example Aspose.Cells .NET | Convert CSV to HTML after applying date filter
// Developer Intent: Load a CSV file, keep rows that match a target date, and generate an HTML file containing only the filtered records.
// Use Cases: Create a daily sales snapshot in HTML from a CSV export. | Provide a web‑ready view of log entries for a particular day. | Produce a date‑specific audit report without manual spreadsheet editing.
// AI Prompts: Generate C# code that filters CSV rows by a date range with Aspose.Cells and saves the output as HTML. | Explain how to extend the AutoFilter to include multiple dates and then export the workbook to PDF. | Show how to style the HTML output (fonts, colors) after applying a date filter using Aspose.Cells.

using System;
using Aspose.Cells;

// This example shows how to load a CSV file into an Aspose.Cells workbook, apply an AutoFilter that keeps only rows matching a given DateTime (using Day grouping), and save the filtered result as an HTML page with all data exported via HtmlSaveOptions.
class CsvDateFilterToHtml
{
    static void Main()
    {
        // Paths for input CSV and output HTML
        string csvPath = "input.csv";
        string htmlPath = "output.html";

        // Date to filter (example: keep rows with this exact date)
        DateTime filterDate = new DateTime(2023, 1, 15);

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Import CSV data (comma‑separated, convert numeric data)
        cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Determine the used range of the first column (assumed to contain dates)
        int lastRow = cells.MaxDataRow;                     // zero‑based index of the last row with data
        sheet.AutoFilter.Range = $"A1:A{lastRow + 1}";      // set autofilter range (including header)

        // Apply a date filter on the first column (field index 0)
        // Grouping by Day and specify year, month, day; time parts are set to zero
        sheet.AutoFilter.AddDateFilter(
            fieldIndex: 0,
            dateTimeGroupingType: DateTimeGroupingType.Day,
            year: filterDate.Year,
            month: filterDate.Month,
            day: filterDate.Day,
            hour: 0,
            minute: 0,
            second: 0);

        // Prepare HTML save options – export all data
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportDataOptions = HtmlExportDataOptions.All;

        // Save the filtered workbook as an HTML file
        workbook.Save(htmlPath, htmlOptions);
    }
}
