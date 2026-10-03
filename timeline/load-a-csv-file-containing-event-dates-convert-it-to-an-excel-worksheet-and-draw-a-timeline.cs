// Title: Generate an Excel event timeline scatter chart from a CSV file using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a CSV containing event names and dates into an Aspose.Cells workbook, adds a sequential index column, and builds a scatter chart where the dates are the X‑axis and the index is the Y‑axis, showing each event name as a data label. | Adapt the example to use a line chart with markers instead of a scatter chart while preserving the CSV source, index column, and event‑name labels.
// Common Searches: how to create a timeline chart from CSV data using Aspose.Cells in C# | Aspose.Cells scatter chart with dates on X axis and custom labels | C# load CSV into workbook and add helper column for chart positioning | export event timeline to Excel file using Aspose.Cells .NET library
// Tags: load csv with Aspose.Cells LoadOptions | add sequential index column for chart Y axis | scatter chart with date X values Aspose.Cells | data labels show category name Aspose.Cells chart | save workbook as Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The code loads events.csv, inserts an index column, creates a scatter chart that maps dates (column B) to the X‑axis and the index (column C) to the Y‑axis, displays event names from column A as data labels, and saves the result as EventTimeline.xlsx.
class TimelineGenerator
{
    static void Main()
    {
        try
        {
            const string inputFile = "events.csv";
            const string outputFile = "EventTimeline.xlsx";

            // Verify that the input CSV file exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the CSV file into a workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Get the first worksheet (the CSV data is placed here)
            Worksheet sheet = workbook.Worksheets[0];

            // Assume the CSV has two columns: A = Event Name, B = Event Date
            // Add a helper column C that contains a sequential index (used for Y‑axis positioning)
            int totalRows = sheet.Cells.MaxDataRow + 1; // includes header row
            sheet.Cells[0, 2].PutValue("Index"); // header for column C
            for (int row = 1; row < totalRows; row++)
            {
                sheet.Cells[row, 2].PutValue(row); // simple 1‑based index
            }

            // Create a scatter chart (X‑axis = Date, Y‑axis = Index)
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Event Timeline";

            // Add the Y‑values series (the index column)
            int seriesIdx = chart.NSeries.Add($"C2:C{totalRows}", true);
            // Assign X‑values (the date column) to the series
            chart.NSeries[seriesIdx].XValues = $"B2:B{totalRows}";

            // Configure data labels (show event names from column A)
            chart.NSeries[seriesIdx].DataLabels.ShowValue = false;
            chart.NSeries[seriesIdx].DataLabels.ShowCategoryName = true; // show category (event name)
            chart.NSeries[seriesIdx].DataLabels.ShowSeriesName = false;

            // Save the result as an Excel workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Timeline chart saved to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
