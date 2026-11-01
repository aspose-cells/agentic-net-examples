// Title: How to log each PivotTable's RefreshDate from all worksheets in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Create a C# console program that opens an .xlsx file with Aspose.Cells, iterates over every worksheet, reads the RefreshDate of each PivotTable, and writes a line with the sheet name, pivot name, and date to a text log. | Generate .NET code that verifies a workbook exists, enumerates all PivotTables, formats the nullable RefreshDate as ISO‑8601 or "N/A", and saves the results to a report file.
// Common Searches: asp.net retrieve pivot table refresh date using Aspose.Cells | c# write pivot table refresh timestamps to a log file | enumerate all pivot tables in an Excel workbook with Aspose.Cells .NET | how to export pivot table metadata from Excel using Aspose.Cells | handle nullable RefreshDate property of PivotTable in C#
// Tags: Aspose.Cells extract pivot refresh date | C# enumerate pivot tables in workbook | log pivot metadata to text file | Aspose.Cells iterate worksheets | handle nullable RefreshDate in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads 'input.xlsx' with Aspose.Cells, checks the file's existence, loops through each worksheet and its PivotTables, reads the nullable RefreshDate, formats it as ISO‑8601 or "N/A", and writes a line containing the worksheet name, pivot name, and refresh date to 'PivotRefreshLog.txt', with basic exception handling.
class PivotTableRefreshLogger
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string workbookPath = "input.xlsx";

            // Verify that the workbook file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Path to the log file
            string logFilePath = "PivotRefreshLog.txt";

            // Write pivot table refresh information to the log file
            using (StreamWriter writer = new StreamWriter(logFilePath, false))
            {
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pivot in sheet.PivotTables)
                    {
                        // Retrieve the RefreshDate (nullable DateTime)
                        DateTime? refreshDate = pivot.RefreshDate;

                        // Format the date or indicate N/A
                        string dateText = refreshDate.HasValue
                            ? refreshDate.Value.ToString("o")
                            : "N/A";

                        writer.WriteLine($"{sheet.Name} - PivotTable '{pivot.Name}' RefreshDate: {dateText}");
                    }
                }
            }

            // Optional: save the workbook if modifications are made
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
