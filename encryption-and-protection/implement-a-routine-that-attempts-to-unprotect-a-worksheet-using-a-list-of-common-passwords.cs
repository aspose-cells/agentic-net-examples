// Title: Unprotect an Aspose.Cells Worksheet by Trying Common Passwords (C#)
// Description: Creates a workbook, protects the first worksheet with a known password, then iterates through a predefined list of common passwords using worksheet.Unprotect until the sheet is no longer protected, and finally saves the file.
// Keywords: Aspose.Cells | C# | worksheet unprotect | common passwords | password list | protect worksheet | worksheet.Unprotect | Excel protection removal | Aspose.Cells API example | brute‑force worksheet password
// Common Searches: how to unprotect an Aspose.Cells worksheet in C# | C# code to try multiple passwords for Excel sheet protection | Aspose.Cells unprotect worksheet with password list | remove worksheet protection programmatically Aspose.Cells | iterate over common passwords to unlock Excel sheet using Aspose
// Developer Intent: Programmatically test a series of likely passwords to remove worksheet protection when the original password is unknown.
// Use Cases: Recover a locked worksheet when the password may be a common or default value. | Prepare workbooks for automated data extraction by clearing unknown protections in bulk. | Log each failed attempt and capture the successful password for compliance reporting.
// AI Prompts: Write C# code that loops through a collection of passwords, calls worksheet.Unprotect for each, stops when worksheet.IsProtected becomes false, and includes exception handling. | Provide an Aspose.Cells example that records every unsuccessful password attempt and outputs the password that successfully unprotects the worksheet. | Suggest an efficient approach for handling a large password list (e.g., parallel processing) while ensuring the workbook remains consistent and thread‑safe.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, protects the first worksheet with a known password, then iterates through a predefined list of common passwords using worksheet.Unprotect until the sheet is no longer protected, and finally saves the file.
    public class WorksheetUnprotectWithCommonPasswords
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Protect the worksheet with a known password
                string actualPassword = "secret";
                worksheet.Protect(ProtectionType.All, actualPassword, null);
                Console.WriteLine("Worksheet protected: " + worksheet.IsProtected);

                // List of common passwords to try
                List<string> commonPasswords = new List<string>
                {
                    "123",
                    "password",
                    "admin",
                    "test",
                    "password123",
                    "secret",          // correct password included for demonstration
                    "letmein"
                };

                // Attempt to unprotect using each password
                foreach (string pwd in commonPasswords)
                {
                    try
                    {
                        // Try to unprotect with the current password
                        worksheet.Unprotect(pwd);

                        // If no exception, check if unprotected
                        if (!worksheet.IsProtected)
                        {
                            Console.WriteLine($"Successfully unprotected worksheet with password: \"{pwd}\"");
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Unprotect failed, continue to next password
                        Console.WriteLine($"Password \"{pwd}\" failed: {ex.Message}");
                    }
                }

                // Final status
                Console.WriteLine("Final worksheet protected state: " + worksheet.IsProtected);

                // Save the workbook
                string outputPath = "WorksheetUnprotected.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
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
            WorksheetUnprotectWithCommonPasswords.Run();
        }
    }
}
