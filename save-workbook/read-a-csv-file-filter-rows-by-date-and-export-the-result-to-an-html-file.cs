// Title: Filter CSV rows by a specific date and save as HTML using Aspose.Cells for .NET (C#)
// AI Prompts: Read a CSV file with Aspose.Cells, keep only rows where the date column is on or after 2023‑01‑01, and write the filtered data to an HTML file. | Create a C# console app that loads a CSV into a Workbook, copies the header and qualifying rows to a new worksheet, and exports the new workbook as HTML using HtmlSaveOptions.
// Common Searches: Aspose.Cells C# filter CSV rows by date and export to HTML | How to load a CSV into an Aspose.Cells workbook and apply a date cutoff in .NET | Save filtered CSV data as HTML using Aspose.Cells SaveOptions | C# example for copying rows between worksheets in Aspose.Cells | Export Aspose.Cells workbook to HTML after filtering data
// Tags: date-based CSV filtering Aspose.Cells | HTML export of filtered workbook Aspose.Cells | load CSV into Aspose.Cells workbook C# | worksheet row copy Aspose.Cells | Aspose.Cells date column filter .NET

using System;
using System.IO;
using Aspose.Cells;

// The program loads a CSV file into an Aspose.Cells Workbook, copies the header and rows whose date column meets a specified cutoff into a new workbook, and saves the filtered workbook as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            // Input CSV file path
            string csvPath = @"C:\Data\input.csv";

            // Output HTML file path
            string htmlPath = @"C:\Data\output.html";

            // Verify that the input CSV exists
            if (!File.Exists(csvPath))
                throw new FileNotFoundException($"Input CSV file not found: {csvPath}");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(htmlPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Define the date filter (e.g., keep rows where the date column is on or after this date)
            DateTime filterDate = new DateTime(2023, 1, 1);

            // Load the CSV file into a workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook sourceWorkbook = new Workbook(csvPath, loadOptions);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Create a new workbook to hold filtered data
            Workbook filteredWorkbook = new Workbook();
            Worksheet filteredSheet = filteredWorkbook.Worksheets[0];
            filteredSheet.Name = "FilteredData";

            // Assume the first row contains headers
            int headerRowIndex = 0;
            int dateColumnIndex = 0; // Change this to the zero‑based index of the date column in the CSV

            // Copy header row to the filtered sheet
            for (int col = 0; col <= sourceSheet.Cells.MaxDataColumn; col++)
            {
                string headerValue = sourceSheet.Cells[headerRowIndex, col].StringValue;
                filteredSheet.Cells[0, col].PutValue(headerValue);
            }

            int filteredRowIndex = 1; // Start after header

            // Iterate through data rows and apply the date filter
            for (int row = headerRowIndex + 1; row <= sourceSheet.Cells.MaxDataRow; row++)
            {
                Cell dateCell = sourceSheet.Cells[row, dateColumnIndex];
                if (dateCell == null || string.IsNullOrWhiteSpace(dateCell.StringValue))
                    continue; // Skip empty rows

                // Try to parse the date
                if (DateTime.TryParse(dateCell.StringValue, out DateTime rowDate))
                {
                    // Keep the row if it meets the filter condition
                    if (rowDate >= filterDate)
                    {
                        // Copy the entire row to the filtered sheet
                        for (int col = 0; col <= sourceSheet.Cells.MaxDataColumn; col++)
                        {
                            Cell srcCell = sourceSheet.Cells[row, col];
                            Cell destCell = filteredSheet.Cells[filteredRowIndex, col];
                            destCell.PutValue(srcCell.Value);
                        }
                        filteredRowIndex++;
                    }
                }
            }

            // Save the filtered workbook as HTML
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            filteredWorkbook.Save(htmlPath, htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
