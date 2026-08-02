// Title: Unprotect an Excel worksheet with Aspose.Cells for .NET using an environment‑variable password
// Description: Loads a protected workbook, reads the worksheet password from the WORKSHEET_PASSWORD environment variable (fallback to empty string), calls Worksheet.Unprotect, and saves the unprotected file. Includes file‑existence checks and exception handling for secure, automated processing.
// Keywords: Aspose.Cells unprotect worksheet .NET | C# read password from environment variable | Excel sheet protection removal programmatically | secure worksheet unprotection | load protected workbook Aspose.Cells
// Common Searches: Aspose.Cells unprotect worksheet using env variable | C# remove Excel sheet password without hard‑coding | how to unprotect Excel worksheet in .NET securely | read worksheet password from environment variable C#
// Developer Intent: Remove worksheet protection by supplying a password obtained from an environment variable.
// Use Cases: Automate batch unprotection of Excel files in CI/CD pipelines where the password is stored as a secret environment variable. | Create a secure utility that unlocks a specific worksheet before data extraction or transformation. | Integrate worksheet unprotection into a serverless function that reads the password from the runtime environment and outputs an unprotected copy.
// AI Prompts: Generate C# code with Aspose.Cells that unprotects a worksheet using a password read from an environment variable and includes robust error handling. | Modify the example to iterate over all worksheets, unprotect each with the same environment‑variable password, and log the operation results. | Write a PowerShell script that sets the WORKSHEET_PASSWORD environment variable, runs the compiled .NET program, and verifies that the output workbook is unprotected.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a protected workbook, reads the worksheet password from the WORKSHEET_PASSWORD environment variable (fallback to empty string), calls Worksheet.Unprotect, and saves the unprotected file. Includes file‑existence checks and exception handling for secure, automated processing.
    public class WorksheetUnprotectWithEnvPassword
    {
        public static void Run()
        {
            try
            {
                // Retrieve the worksheet password from an environment variable.
                // If not set, use an empty string (unprotects worksheets without a password).
                string password = Environment.GetEnvironmentVariable("WORKSHEET_PASSWORD") ?? string.Empty;

                string inputPath = "protected.xlsx";
                string outputPath = "unprotected.xlsx";

                // Ensure the input workbook exists before loading.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the protected workbook.
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (or adjust index as needed).
                Worksheet worksheet = workbook.Worksheets[0];

                // Unprotect the worksheet using the retrieved password.
                worksheet.Unprotect(password);

                // Save the workbook after unprotecting the worksheet.
                workbook.Save(outputPath);
                Console.WriteLine($"Worksheet unprotected and saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point.
    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetUnprotectWithEnvPassword.Run();
        }
    }
}
