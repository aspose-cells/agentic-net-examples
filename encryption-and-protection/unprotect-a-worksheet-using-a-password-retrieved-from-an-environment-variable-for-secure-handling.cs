using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetUnprotect
{
    class Program
    {
        static void Main()
        {
            // Path to the protected workbook
            string inputPath = "protected.xlsx";

            // Load the workbook (no password needed for opening the file itself)
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the worksheet protection password from an environment variable
            // Ensure the environment variable WS_PASSWORD is set securely in the execution environment
            string worksheetPassword = Environment.GetEnvironmentVariable("WS_PASSWORD");

            if (string.IsNullOrEmpty(worksheetPassword))
            {
                Console.WriteLine("Environment variable 'WS_PASSWORD' is not set.");
                return;
            }

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Unprotect the worksheet using the retrieved password
            worksheet.Unprotect(worksheetPassword);

            // Verify that the worksheet is no longer protected
            Console.WriteLine($"Worksheet protected status after unprotect: {worksheet.IsProtected}");

            // Save the unprotected workbook to a new file
            string outputPath = "unprotected.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}