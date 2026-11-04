// Title: Unprotect an Excel worksheet in C# with Aspose.Cells using a password from an environment variable
// Common Searches: C# Aspose.Cells unprotect worksheet using environment variable password | How to read sheet password from ENV and remove protection with Aspose.Cells | Load password‑protected Excel file and unprotect a sheet in .NET | Aspose.Cells unprotect sheet when workbook also has a password | Retrieve worksheet password from system environment variable in C#
// Tags: worksheet unprotection Aspose.Cells C# | environment variable password retrieval .NET | load workbook with file password Aspose.Cells | remove sheet protection programmatically | secure Excel password handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook (optionally using a file password), obtains the worksheet password from the SHEET_PASSWORD environment variable, unprotects the first worksheet with Aspose.Cells, and saves the result to a new file, while providing error handling for missing files or variables.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the Excel file to be processed
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Retrieve passwords from environment variables
            string workbookPassword = Environment.GetEnvironmentVariable("WORKBOOK_PASSWORD"); // optional file password
            string sheetPassword = Environment.GetEnvironmentVariable("SHEET_PASSWORD");       // worksheet password

            if (string.IsNullOrEmpty(sheetPassword))
            {
                Console.WriteLine("Error: Environment variable 'SHEET_PASSWORD' is not set.");
                return;
            }

            // Load the workbook, providing the file password if required
            Workbook workbook;
            if (!string.IsNullOrEmpty(workbookPassword))
            {
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = workbookPassword
                };
                workbook = new Workbook(inputPath, loadOptions);
            }
            else
            {
                workbook = new Workbook(inputPath);
            }

            // Unprotect the first worksheet using the worksheet password
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Unprotect(sheetPassword);

            // Save the modified workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Worksheet unprotected and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
