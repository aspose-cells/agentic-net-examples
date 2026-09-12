// Title: Set a query table’s RefreshPeriod to 30 minutes and enable RefreshOnFileOpen using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that sets the RefreshPeriod of the first worksheet’s query table to 30 minutes and turns on RefreshOnFileOpen. | Show how to detect a query table, change its automatic refresh interval, and save the workbook programmatically with Aspose.Cells.
// Common Searches: Aspose.Cells C# set query table refresh period to 30 minutes | enable automatic refresh on file open for Excel query tables using Aspose.Cells | programmatically update query table properties in a .NET workbook | how to change RefreshPeriod property of a QueryTable with Aspose.Cells | refresh query table every half hour Aspose.Cells example
// Tags: Aspose.Cells query table automatic refresh | C# enable query table refresh on open | modify query table settings .NET | Excel query table refresh interval Aspose.Cells | set query table RefreshPeriod programmatically

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// Loads an Excel workbook, checks for a query table in the first worksheet, sets its RefreshPeriod to 30 minutes, enables RefreshOnFileOpen, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing a query table
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one query table
            if (sheet.QueryTables.Count > 0)
            {
                // Get the first query table
                QueryTable queryTable = sheet.QueryTables[0];
                Type qtType = queryTable.GetType();

                // Set refresh interval (if the property exists)
                PropertyInfo refreshPeriodProp = qtType.GetProperty("RefreshPeriod");
                if (refreshPeriodProp != null && refreshPeriodProp.CanWrite)
                {
                    refreshPeriodProp.SetValue(queryTable, 30);
                }

                // Enable refresh on file open (if the property exists)
                PropertyInfo refreshOnOpenProp = qtType.GetProperty("RefreshOnFileOpen");
                if (refreshOnOpenProp != null && refreshOnOpenProp.CanWrite)
                {
                    refreshOnOpenProp.SetValue(queryTable, true);
                }
            }
            else
            {
                Console.WriteLine("No query tables found in the first worksheet.");
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
