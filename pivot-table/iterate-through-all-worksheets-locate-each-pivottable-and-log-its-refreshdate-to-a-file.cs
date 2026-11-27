// Title: C# – Log the RefreshDate of Each PivotTable in All Worksheets to a Text File with Aspose.Cells
// Common Searches: how to write pivot table refresh dates to a log file using Aspose.Cells C# | Aspose.Cells iterate all worksheets and export pivot RefreshDate | C# code to retrieve and save PivotTable RefreshDate from an Excel workbook | log pivot table last refreshed timestamp for each sheet with Aspose.Cells | extract pivot refresh metadata from multiple worksheets in .NET
// Tags: Aspose.Cells pivot table RefreshDate extraction | log pivot metadata to text file C# | iterate worksheets to access PivotTables .NET | export pivot refresh timestamps using Aspose.Cells | C# write Excel pivot information to log

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, loops through every worksheet and each PivotTable it contains, reads the PivotTable's RefreshDate (or marks it as "Never" when unavailable), and writes the worksheet name, pivot name, and refresh date to a specified text log file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string logPath = "PivotRefreshLog.txt";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Create or overwrite the log file
            using (StreamWriter writer = new StreamWriter(logPath, false))
            {
                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all pivot tables on the worksheet
                    foreach (PivotTable pivot in sheet.PivotTables)
                    {
                        // Determine refresh information
                        string refreshInfo = pivot.RefreshDate != DateTime.MinValue
                            ? pivot.RefreshDate.ToString("o") // ISO 8601 format
                            : "Never";

                        // Write details to the log
                        writer.WriteLine($"Worksheet: {sheet.Name}, PivotTable: {pivot.Name}, RefreshDate: {refreshInfo}");
                    }
                }
            }

            Console.WriteLine($"Pivot refresh log written to {logPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
