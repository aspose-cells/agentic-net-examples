// Title: C# – Open a password‑protected XLS file and retrieve the chart on the second worksheet with Aspose.Cells
// Description: This example demonstrates how to load a protected XLS workbook using Aspose.Cells LoadOptions, access the worksheet at index 1, verify the presence of charts, and extract the first chart’s type and parent sheet name. It also includes file‑existence checks and exception handling for invalid passwords or missing files.
// Keywords: Aspose.Cells password protected XLS | load encrypted Excel file C# | retrieve chart from worksheet | access charts in protected workbook | Aspose.Cells LoadOptions example
// Common Searches: open password protected xls with Aspose.Cells | get chart from second sheet in protected Excel file | C# read chart type from encrypted workbook | Aspose.Cells chart extraction from secured file
// Developer Intent: Open a secured XLS workbook and obtain the chart located on its second worksheet.
// Use Cases: Automated reporting that reads chart metadata from confidential Excel files. | Pre‑processing validation to ensure a specific sheet contains at least one chart before further analysis. | Generating documentation that lists chart types and their source worksheets from protected workbooks.
// AI Prompts: Generate C# code using Aspose.Cells to open a password‑protected XLS file and list all charts on a given worksheet index. | Show how to handle incorrect passwords and missing files when loading an encrypted Excel workbook with Aspose.Cells. | Provide a loop that iterates through every chart on the second worksheet of a protected workbook and logs each chart’s type and title.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example demonstrates how to load a protected XLS workbook using Aspose.Cells LoadOptions, access the worksheet at index 1, verify the presence of charts, and extract the first chart’s type and parent sheet name. It also includes file‑existence checks and exception handling for invalid passwords or missing files.
class RetrieveChartFromProtectedWorkbook
{
    static void Main()
    {
        // Path to the password‑protected XLS file
        string filePath = "protected.xls";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file '{filePath}' was not found.");
            return;
        }

        try
        {
            // Set the password required to open the workbook
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "yourPassword"
            };

            // Open the workbook using the load options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Access the second worksheet (index 1)
            Worksheet secondSheet = workbook.Worksheets[1];

            // Check if the worksheet contains any charts
            if (secondSheet.Charts.Count > 0)
            {
                // Retrieve the first chart on the second worksheet
                Chart chart = secondSheet.Charts[0];

                // Demonstrate accessing chart properties
                Console.WriteLine("Chart type: " + chart.Type);
                Console.WriteLine("Chart belongs to worksheet: " + chart.Worksheet.Name);
            }
            else
            {
                Console.WriteLine("No charts found on the second worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
