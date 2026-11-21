// Title: Update an Excel chart’s data source dynamically for a user‑selected date range using Aspose.Cells for .NET
// Common Searches: aspnet change chart series range based on start and end dates with Aspose.Cells | c# programmatically set chart XValues to a filtered date range in Excel using Aspose.Cells | how to bind an Excel chart to rows that match a date interval in Aspose.Cells .NET | update Excel chart data source after selecting a date range in a C# application | Aspose.Cells dynamic chart data range from worksheet based on date column
// Tags: Aspose.Cells modify chart series data | C# filter worksheet by date column | Excel chart XValues assignment using Aspose.Cells | dynamic chart data source in .NET | date‑based chart series selection Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads a workbook, scans the 'Data' sheet for rows whose dates in column A fall within a specified start‑end interval, builds A1‑style range strings for those rows, assigns the ranges to the XValues and Values of the first series of the first chart, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Worksheet that contains the source data (assumed name "Data")
            Worksheet dataSheet = workbook.Worksheets["Data"];
            if (dataSheet == null)
            {
                Console.WriteLine("Worksheet named 'Data' was not found.");
                return;
            }

            Cells cells = dataSheet.Cells;

            // User‑selected date range (example values; replace with actual input)
            DateTime startDate = new DateTime(2023, 1, 1);
            DateTime endDate   = new DateTime(2023, 12, 31);

            // Determine the first and last rows that fall within the date range
            int firstRow = -1;
            int lastRow  = -1;
            int maxRow   = cells.MaxDataRow; // last row that contains data

            for (int row = 0; row <= maxRow; row++)
            {
                object cellValue = cells[row, 0].Value; // Column A holds dates
                if (cellValue is DateTime dt)
                {
                    if (dt >= startDate && dt <= endDate)
                    {
                        if (firstRow == -1) firstRow = row;   // mark start
                        lastRow = row;                        // keep updating end
                    }
                }
            }

            // If no rows match the criteria, exit
            if (firstRow == -1)
            {
                Console.WriteLine("No data found for the selected date range.");
                return;
            }

            // Build the range strings for X (dates) and Y (values) axes (1‑based indexing)
            string sheetName = dataSheet.Name;
            string xRange = $"'{sheetName}'!A{firstRow + 1}:A{lastRow + 1}";
            string yRange = $"'{sheetName}'!B{firstRow + 1}:B{lastRow + 1}";

            // Access the chart (assumed first chart on the first worksheet)
            Worksheet chartSheet = workbook.Worksheets[0];
            if (chartSheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            Chart chart = chartSheet.Charts[0];

            // Update the chart series to use the new dynamic ranges
            if (chart.NSeries.Count > 0)
            {
                chart.NSeries[0].XValues = xRange; // dates on X‑axis
                chart.NSeries[0].Values  = yRange; // corresponding values
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
