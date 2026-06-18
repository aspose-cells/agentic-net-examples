using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

class SmartMarkerFilterDemo
{
    static void Main()
    {
        try
        {
            const string templatePath = "SmartMarkerTemplate.xlsx";
            const string outputPath = "FilteredSmartMarkers.xlsx";

            // Verify that the template file exists
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load the workbook that contains smart markers
            Workbook workbook = new Workbook(templatePath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample data source for the smart markers
            List<Person> persons = new List<Person>
            {
                new Person { Name = "John", Age = 30, Country = "US" },
                new Person { Name = "Anna", Age = 25, Country = "UK" },
                new Person { Name = "Mike", Age = 35, Country = "US" }
            };

            // Process the smart markers with the data source
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Persons", persons);
            designer.Process();

            // Determine the data range after processing
            int lastDataRow = worksheet.Cells.MaxDataRow;          // zero‑based
            int lastDataColumn = worksheet.Cells.MaxDataColumn;    // zero‑based

            // Build the range string (e.g., "A1:C10")
            string range = $"A1:{CellsHelper.CellIndexToName(lastDataColumn, lastDataRow)}";
            worksheet.AutoFilter.Range = range;

            // Find the column index of the "Country" header (case‑insensitive)
            int countryColIndex = -1;
            for (int col = 0; col <= lastDataColumn; col++)
            {
                var header = worksheet.Cells[0, col].StringValue;
                if (string.Equals(header, "Country", StringComparison.OrdinalIgnoreCase))
                {
                    countryColIndex = col;
                    break;
                }
            }

            if (countryColIndex == -1)
                throw new InvalidOperationException("Country column not found in the header row.");

            // Apply filter to show only rows where Country = "US"
            worksheet.AutoFilter.AddFilter(countryColIndex, "US");
            worksheet.AutoFilter.Refresh();

            // Auto‑fit rows while ignoring hidden rows
            AutoFitterOptions options = new AutoFitterOptions
            {
                IgnoreHidden = true,
                AutoFitMergedCells = false,
                OnlyAuto = false
            };
            worksheet.AutoFitRows(options);

            // Save the resulting workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simple POCO class representing the data source
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Country { get; set; }
    }
}