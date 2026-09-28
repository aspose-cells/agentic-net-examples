// Title: Unprotect an Excel worksheet using a password stored in an environment variable with Aspose.Cells for .NET
// AI Prompts: Write C# code that reads the worksheet password from the WORKSHEET_PASSWORD environment variable and calls Worksheet.Unprotect in Aspose.Cells. | Show how to fall back to Worksheet.Unprotect() when the environment variable is missing or empty. | Demonstrate unprotecting multiple worksheets where each sheet's password is retrieved from a separate environment variable. | Explain how to integrate environment‑based password retrieval into an existing Aspose.Cells workbook processing pipeline.
// Common Searches: asp.net unprotect Excel sheet using password from environment variable Aspose.Cells | c# read env variable for worksheet password and remove protection with Aspose.Cells | how to programmatically unprotect a protected worksheet when password is stored in OS environment variable | Aspose.Cells unprotect sheet without hard‑coding password in code | example of conditional worksheet unprotection based on environment variable in C#
// Tags: worksheet unprotect with environment variable Aspose.Cells | Aspose.Cells read password from env in .NET | C# remove Excel sheet protection programmatically | secure worksheet unprotection using env variable | load workbook and unprotect sheet Aspose.Cells | conditional unprotect worksheet C#

using System;
using Aspose.Cells;

// // Loads "input.xlsx", obtains the worksheet password from the WORKSHEET_PASSWORD environment variable, unprotects the first worksheet (using the password if present or falling back to password‑less unprotect), and saves the result as "output.xlsx".
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Retrieve the password from an environment variable for secure handling
        string password = Environment.GetEnvironmentVariable("WORKSHEET_PASSWORD");

        // Unprotect the worksheet using the retrieved password
        if (!string.IsNullOrEmpty(password))
        {
            sheet.Unprotect(password);
        }
        else
        {
            // If no password is set, attempt to unprotect without a password
            sheet.Unprotect();
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
