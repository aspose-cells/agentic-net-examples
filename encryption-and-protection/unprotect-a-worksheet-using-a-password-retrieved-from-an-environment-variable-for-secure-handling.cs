// Title: Unprotect an Excel worksheet with Aspose.Cells (.NET) using an environment‑variable password
// Description: Loads "protected.xlsx", reads the worksheet password from the WORKSHEET_PASSWORD environment variable, unprotects the first sheet, handles missing files or wrong passwords, and saves the result as "unprotected.xlsx".
// Keywords: Aspose.Cells | C# | .NET | unprotect worksheet | environment variable password | Excel protection removal | secure Excel automation | workbook unprotect code | password retrieval from env | Excel security best practices
// Common Searches: Aspose.Cells unprotect worksheet using env variable | C# read password from environment variable to unprotect Excel sheet | how to remove worksheet protection with Aspose.Cells .NET | unprotect Excel file programmatically without hard‑coding password | error handling for invalid worksheet password Aspose.Cells
// Developer Intent: Remove worksheet protection in a .NET application by supplying a password obtained securely from an environment variable.
// Use Cases: Automated ETL pipeline that receives password‑protected Excel files and must unlock them before data extraction. | CI/CD job that temporarily unprotects test workbooks using a secret stored in the build environment. | Desktop utility that opens user‑provided protected workbooks, unprotects the first sheet with a system‑wide password, and saves an unprotected copy.
// AI Prompts: Write C# code with Aspose.Cells to unprotect the first worksheet of a workbook, reading the password from an environment variable and handling missing files or wrong passwords. | Show how to fetch the worksheet password from Azure Key Vault and pass it to Aspose.Cells for unprotecting a sheet in .NET. | Explain how to extend the example to loop through all worksheets and unprotect each using separate environment‑variable passwords.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUnprotectExample
{
    // Loads "protected.xlsx", reads the worksheet password from the WORKSHEET_PASSWORD environment variable, unprotects the first sheet, handles missing files or wrong passwords, and saves the result as "unprotected.xlsx".
    class Program
    {
        static void Main()
        {
            try
            {
                // Retrieve the worksheet password from an environment variable
                string worksheetPassword = Environment.GetEnvironmentVariable("WORKSHEET_PASSWORD") ?? string.Empty;

                const string inputPath = "protected.xlsx";
                const string outputPath = "unprotected.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook that contains the protected worksheet
                Workbook workbook = new Workbook(inputPath);

                // Unprotect the first worksheet if a password is provided
                Worksheet sheet = workbook.Worksheets[0];
                if (!string.IsNullOrEmpty(worksheetPassword))
                {
                    try
                    {
                        sheet.Unprotect(worksheetPassword);
                    }
                    catch (CellsException ex)
                    {
                        // Handle invalid password or other unprotect errors
                        Console.WriteLine($"Error while unprotecting worksheet: {ex.Message}");
                        return;
                    }
                }
                else
                {
                    Console.WriteLine("Warning: No worksheet password provided. Skipping unprotect step.");
                }

                // Save the workbook with the worksheet now unprotected
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
