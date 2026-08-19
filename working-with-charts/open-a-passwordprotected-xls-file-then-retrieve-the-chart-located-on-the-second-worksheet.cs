// Title: C# – Open a password‑protected XLS workbook and retrieve the first chart from the second worksheet with Aspose.Cells
// Description: Loads a password‑protected XLS file using Aspose.Cells LoadOptions, accesses the worksheet at index 1, checks for charts, and returns the first chart’s type and its parent sheet name.
// Keywords: Aspose.Cells C# load password protected XLS | retrieve chart from second worksheet Aspose.Cells | open protected Excel file .NET | Aspose.Cells chart extraction | LoadOptions.Password example | read chart from protected workbook | C# Excel chart access | Aspose.Cells for .NET chart API
// Common Searches: how to open a password protected Excel file with Aspose.Cells | Aspose.Cells get chart from second sheet in protected workbook | C# read chart from protected XLS using Aspose.Cells | Aspose.Cells load options password example | retrieve chart type from protected Excel file
// Developer Intent: Open a password‑protected XLS workbook and obtain the chart located on the second worksheet.
// Use Cases: Extract chart metadata (type, sheet name) from a secured workbook for reporting. | Validate that a protected workbook contains a chart on a specific sheet before further processing. | Copy or export the retrieved chart after unlocking the workbook programmatically.
// AI Prompts: Generate C# code that opens a password‑protected XLS file with Aspose.Cells and lists all charts on the second worksheet. | Show how to handle exceptions when loading a protected workbook and retrieving its first chart using Aspose.Cells. | Provide an example of copying the retrieved chart to another workbook after opening a password‑protected file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads a password‑protected XLS file using Aspose.Cells LoadOptions, accesses the worksheet at index 1, checks for charts, and returns the first chart’s type and its parent sheet name.
class RetrieveChartFromProtectedWorkbook
{
    static void Main()
    {
        // Path to the password‑protected XLS file
        string filePath = "protected_workbook.xls";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // Password used to protect the workbook
        string password = "yourPassword";

        try
        {
            // Load the workbook with the specified password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Access the second worksheet (index 1)
            Worksheet secondSheet = workbook.Worksheets[1];

            // Ensure the worksheet contains at least one chart
            if (secondSheet.Charts.Count > 0)
            {
                // Retrieve the first chart on the second worksheet
                Chart chart = secondSheet.Charts[0];

                // Output some chart information
                Console.WriteLine("Chart Type: " + chart.Type);
                Console.WriteLine("Chart is on worksheet: " + chart.Worksheet.Name);
            }
            else
            {
                Console.WriteLine("No charts found on the second worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during loading or processing
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
