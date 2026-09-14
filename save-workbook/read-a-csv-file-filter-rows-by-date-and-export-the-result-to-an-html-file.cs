// Title: Read a CSV with Aspose.Cells, filter rows by a specific date using AutoFilter, and export the filtered data to HTML in C#
// AI Prompts: Generate C# code that loads a CSV file into an Aspose.Cells workbook, sets an AutoFilter on column A for a given date, refreshes the filter, and saves the visible rows as an HTML file using HtmlSaveOptions. | Provide a C# example that imports data.csv with Aspose.Cells, applies a day‑level date filter on the first column (e.g., 2023‑01‑15), and writes the filtered worksheet to filtered.html while exporting all data.
// Common Searches: Aspose.Cells C# filter CSV rows by date and save as HTML | How to use AutoFilter with date criteria on imported CSV in Aspose.Cells .NET | Export filtered worksheet to HTML using Aspose.Cells HtmlSaveOptions | C# code to import CSV, apply date AutoFilter, and generate HTML report with Aspose.Cells | Set date grouping type Day in Aspose.Cells AutoFilter and export result
// Tags: import csv Aspose.Cells C# | date autofilter Aspose.Cells | htmlsaveoptions export all data Aspose.Cells | filter worksheet rows by date Aspose.Cells | save filtered workbook as html Aspose.Cells

using System;
using Aspose.Cells;

// The program imports a CSV file into an Aspose.Cells workbook, defines an AutoFilter covering the data range, adds a day‑level date filter on the first column, refreshes the filter, and saves the visible rows to an HTML file using HtmlSaveOptions configured to export all data.
class Program
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "data.csv";

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Import CSV data (comma delimiter, convert numeric data)
        cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Determine the used range of the worksheet
        int lastRow = cells.MaxDataRow;      // zero‑based index of the last row with data
        int lastCol = cells.MaxDataColumn;   // zero‑based index of the last column with data

        // Set AutoFilter range to cover the whole data block
        string range = CellsHelper.CellIndexToName(0, 0) + ":" + CellsHelper.CellIndexToName(lastRow, lastCol);
        worksheet.AutoFilter.Range = range;

        // Apply a date filter on the first column (index 0)
        // Example: keep rows where the date equals 2023‑01‑15
        worksheet.AutoFilter.AddDateFilter(
            fieldIndex: 0,
            dateTimeGroupingType: DateTimeGroupingType.Day,
            year: 2023,
            month: 1,
            day: 15,
            hour: 0,
            minute: 0,
            second: 0);

        // Refresh the filter to apply changes
        worksheet.AutoFilter.Refresh();

        // Prepare HTML save options to export all data
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportDataOptions = HtmlExportDataOptions.All;

        // Save the filtered result as an HTML file
        workbook.Save("filtered.html", htmlOptions);
    }
}
