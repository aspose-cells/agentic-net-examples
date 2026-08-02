// Title: Check if an Excel worksheet is password‑protected with Aspose.Cells for .NET before editing
// Description: Loads a workbook, accesses the first worksheet, uses Worksheet.Protection.IsProtectedWithPassword to detect password protection, conditionally writes to cell A1 only when the sheet is unprotected, and saves the result.
// Keywords: Aspose.Cells | C# | worksheet protection | IsProtectedWithPassword | password protected Excel | conditional edit | prevent modifications | Excel security | detect sheet password
// Common Searches: Aspose.Cells check worksheet password protection C# | How to know if Excel sheet is password protected using Aspose.Cells | Skip editing protected worksheet Aspose.Cells | C# detect protected worksheet in Excel | IsProtectedWithPassword example
// Developer Intent: Identify whether a worksheet is secured with a password and perform data modifications only on unprotected sheets.
// Use Cases: Batch‑process multiple workbooks while leaving password‑protected sheets unchanged. | Automated data import that respects existing worksheet security. | Generate reports that modify only editable worksheets. | Log the protection status of each sheet for audit before applying changes.
// AI Prompts: Write C# code using Aspose.Cells that opens an Excel file, iterates through all worksheets, checks Worksheet.Protection.IsProtectedWithPassword, and skips any sheet that is password‑protected. | Show how to throw a custom exception when Worksheet.Protection.IsProtectedWithPassword returns true, using Aspose.Cells for .NET. | Create an example that logs each worksheet's protection state and writes data to cell A1 only when the sheet is not password‑protected.

using System;
using Aspose.Cells;

// Loads a workbook, accesses the first worksheet, uses Worksheet.Protection.IsProtectedWithPassword to detect password protection, conditionally writes to cell A1 only when the sheet is unprotected, and saves the result.
class VerifyWorksheetProtection
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Verify if the worksheet is protected with a password
        bool isPasswordProtected = worksheet.Protection.IsProtectedWithPassword;

        if (isPasswordProtected)
        {
            Console.WriteLine("Worksheet is password protected. No modifications will be made.");
        }
        else
        {
            Console.WriteLine("Worksheet is not password protected. Performing data modifications.");

            // Example modification: write a value to cell A1
            worksheet.Cells["A1"].PutValue("Modified");
        }

        // Save the workbook after the conditional operation
        workbook.Save("output.xlsx");
    }
}
