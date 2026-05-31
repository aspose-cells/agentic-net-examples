using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Path to the password‑protected XLS file
        string filePath = "protected.xls";

        // Password used to protect the workbook
        string password = "yourPassword";

        try
        {
            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {Path.GetFullPath(filePath)}");
                return;
            }

            // Load options with the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };

            // Open the workbook using the load options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Ensure the workbook has at least two worksheets
            if (workbook.Worksheets.Count <= 1)
            {
                Console.WriteLine("The workbook does not contain a second worksheet.");
                return;
            }

            // Access the second worksheet (index 1)
            Worksheet secondWorksheet = workbook.Worksheets[1];

            // Retrieve the first chart on the second worksheet, if any
            if (secondWorksheet.Charts.Count > 0)
            {
                Chart chart = secondWorksheet.Charts[0];

                // Display chart type and its worksheet name
                Console.WriteLine($"Chart Type: {chart.Type}");
                Console.WriteLine($"Chart resides on worksheet: {chart.Worksheet.Name}");
            }
            else
            {
                Console.WriteLine("No charts found on the second worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Handle unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}