// Title: Remove and Re‑apply Worksheet Protection with a Case‑Sensitive Password using Aspose.Cells for .NET
// Description: Demonstrates how to unprotect a worksheet in a new workbook, then protect it again with a new case‑sensitive password using Aspose.Cells for .NET. The example shows default case sensitivity, status checks, and saving the workbook.
// Keywords: Aspose.Cells worksheet unprotect | protect worksheet case sensitive password | C# Aspose.Cells remove protection | worksheet password Aspose.Cells .NET | Excel protection programmatically
// Common Searches: how to unprotect an Excel worksheet with Aspose.Cells | Aspose.Cells case sensitive worksheet password example | remove worksheet protection and set new password C# | Aspose.Cells protect worksheet with custom password
// Developer Intent: The developer needs to clear the current worksheet password and then secure the worksheet again using a new case‑sensitive password.
// Use Cases: Reset worksheet locks before applying updated security policies in generated reports. | Migrate existing Excel files to a new corporate password standard while preserving case sensitivity. | Automate creation of Excel workbooks that enforce strong, case‑sensitive passwords for each worksheet.
// AI Prompts: Write C# code that unprotects a worksheet with a known password and then protects it with a different case‑sensitive password using Aspose.Cells. | Explain Aspose.Cells' handling of case sensitivity for worksheet passwords and show how to verify protection status after re‑applying it. | Provide robust error handling for changing worksheet protection in Aspose.Cells, including exception capture and confirmation of workbook save.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to unprotect a worksheet in a new workbook, then protect it again with a new case‑sensitive password using Aspose.Cells for .NET. The example shows default case sensitivity, status checks, and saving the workbook.
    public class WorksheetPasswordCaseSensitiveDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add some sample data (optional)
                worksheet.Cells["A1"].PutValue("Sample Data");

                // Initially protect the worksheet with a password (e.g., "test")
                string initialPassword = "test";
                worksheet.Protect(ProtectionType.All, initialPassword, null);
                Console.WriteLine("Worksheet initially protected: " + worksheet.IsProtected);

                // Remove the existing protection using the correct password
                worksheet.Unprotect(initialPassword);
                Console.WriteLine("Worksheet unprotected: " + !worksheet.IsProtected);

                // Re‑apply protection with a case‑sensitive password
                // Passwords in Aspose.Cells are case‑sensitive by default
                string newPassword = "CaseSensitivePass123";
                worksheet.Protect(ProtectionType.All, newPassword, null);
                Console.WriteLine("Worksheet re‑protected with case‑sensitive password: " + worksheet.IsProtected);

                // Save the workbook to verify the protection
                string outputPath = "WorksheetPasswordCaseSensitiveDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetPasswordCaseSensitiveDemo.Run();
        }
    }
}
