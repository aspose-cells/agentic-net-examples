// Title: Unprotect an Excel worksheet with Aspose.Cells for .NET using an environment‑variable password (C#)
// Description: C# example that loads a protected workbook with Aspose.Cells, reads the worksheet protection password from the environment variable WORKSHEET_PASSWORD, calls Worksheet.Unprotect, confirms the sheet is no longer protected, and saves the result as an unprotected file.
// Keywords: Aspose.Cells | C# | unprotect worksheet | environment variable password | Excel protection | Worksheet.Unprotect | secure password handling | load workbook | save unprotected workbook | Excel automation .NET
// Common Searches: Aspose.Cells unprotect worksheet C# environment variable | How to read worksheet password from env var in .NET | Remove worksheet protection with Aspose.Cells without hard‑coding password | C# code to unprotect Excel sheet using environment variable | Securely unprotect Excel worksheet in CI/CD pipeline Aspose
// Developer Intent: Read a worksheet protection password from a secure environment variable, unprotect the sheet with Aspose.Cells, and save the workbook without protection.
// Use Cases: CI/CD pipelines that process incoming protected Excel files and need to remove worksheet protection without exposing passwords in source code. | Desktop or web applications that open password‑protected worksheets, retrieve the password from a protected config or environment variable, and then modify the sheet. | Scheduled maintenance scripts that batch‑process multiple workbooks, unprotecting worksheets using a centrally managed environment variable.
// AI Prompts: Generate C# code using Aspose.Cells that reads the worksheet protection password from an environment variable, unprotects the first worksheet, verifies the status, and saves the workbook. | Explain best practices for storing worksheet passwords in .NET (environment variables, secret managers) when using Aspose.Cells to unprotect worksheets. | Show how to handle missing or empty environment variables gracefully in an Aspose.Cells worksheet‑unprotect routine.

using System;
using Aspose.Cells;

// C# example that loads a protected workbook with Aspose.Cells, reads the worksheet protection password from the environment variable WORKSHEET_PASSWORD, calls Worksheet.Unprotect, confirms the sheet is no longer protected, and saves the result as an unprotected file.
class UnprotectWorksheetExample
{
    static void Main()
    {
        // Path to the protected workbook
        string inputPath = "protected.xlsx";

        // Load the workbook (no password needed for worksheet protection)
        Workbook workbook = new Workbook(inputPath);

        // Retrieve the worksheet protection password from an environment variable
        // Ensure the environment variable "WORKSHEET_PASSWORD" is set securely
        string worksheetPassword = Environment.GetEnvironmentVariable("WORKSHEET_PASSWORD");

        if (string.IsNullOrEmpty(worksheetPassword))
        {
            Console.WriteLine("Error: Environment variable 'WORKSHEET_PASSWORD' is not set.");
            return;
        }

        // Access the first worksheet (adjust index as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Unprotect the worksheet using the retrieved password
        sheet.Unprotect(worksheetPassword);

        // Verify that the worksheet is no longer protected
        Console.WriteLine($"Worksheet protected status after unprotect: {sheet.IsProtected}");

        // Save the unprotected workbook
        string outputPath = "unprotected.xlsx";
        workbook.Save(outputPath);

        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}
