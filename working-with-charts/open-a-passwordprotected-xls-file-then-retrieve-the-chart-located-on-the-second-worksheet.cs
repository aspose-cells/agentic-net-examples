using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            // Path to the password‑protected XLS file
            string filePath = "protected.xls";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Password required to open the workbook
            string password = "myPassword";

            // Load the workbook with the specified password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Access the second worksheet (index 1)
            Worksheet secondWorksheet = workbook.Worksheets[1];

            // Check if the worksheet contains any charts
            if (secondWorksheet.Charts.Count > 0)
            {
                // Retrieve the first chart on the second worksheet
                Chart chart = secondWorksheet.Charts[0];

                // Example usage: display chart type and its parent worksheet name
                Console.WriteLine($"Chart Type: {chart.Type}");
                Console.WriteLine($"Chart is located in worksheet: {chart.Worksheet.Name}");
            }
            else
            {
                Console.WriteLine("No charts found on the second worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}