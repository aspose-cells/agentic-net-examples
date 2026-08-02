// Title: Aspose.Cells .NET – Protect a Worksheet with an Empty Password and Check API Results
// Description: Demonstrates how to protect the first worksheet of a new workbook using a blank password, prints the protection flags (IsProtected, IsProtectedWithPassword) and VerifyPassword outcomes for empty and null inputs, saves the file, reloads it, and verifies that the protection state remains consistent.
// Keywords: Aspose.Cells protect worksheet empty password | worksheet IsProtected flag | IsProtectedWithPassword false blank password | VerifyPassword empty string Aspose.Cells | save and reload protected worksheet .NET | C# Aspose.Cells worksheet protection example | blank password worksheet protection behavior
// Common Searches: Aspose.Cells protect worksheet with no password | IsProtectedWithPassword returns false for blank password | VerifyPassword empty string after worksheet protection | Does saving preserve empty‑password protection in Aspose.Cells | C# example protecting worksheet without a password
// Developer Intent: Find out how Aspose.Cells behaves when a worksheet is protected with an empty password and whether the protection persists after the workbook is saved and reopened.
// Use Cases: Confirm that Worksheet.IsProtected is true while Protection.IsProtectedWithPassword is false when the password is empty. | Validate that protection.VerifyPassword("") returns true and protection.VerifyPassword(null) returns false for a blank‑password protection. | Ensure the protection flags and verification results stay unchanged after saving the workbook and loading it again.
// AI Prompts: Explain the effect of using an empty string as the password in Aspose.Cells worksheet protection, including the values of IsProtected, IsProtectedWithPassword, and VerifyPassword before and after saving the file. | Generate C# code that protects a worksheet with a blank password, saves the workbook, reloads it, and asserts the expected protection properties. | Provide best‑practice recommendations for handling scenarios where a worksheet might be unintentionally protected with an empty password in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to protect the first worksheet of a new workbook using a blank password, prints the protection flags (IsProtected, IsProtectedWithPassword) and VerifyPassword outcomes for empty and null inputs, saves the file, reloads it, and verifies that the protection state remains consistent.
    public class WorksheetEmptyPasswordProtectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Access the worksheet protection object
                Protection protection = sheet.Protection;

                // Set an empty password (blank string)
                protection.Password = "";

                // Protect the worksheet using the empty password (provide oldPassword as empty)
                sheet.Protect(ProtectionType.All, "", "");

                // Observe protection status immediately after protecting
                Console.WriteLine("After protection:");
                Console.WriteLine($"Worksheet.IsProtected: {sheet.IsProtected}");
                Console.WriteLine($"Protection.IsProtectedWithPassword: {protection.IsProtectedWithPassword}");
                Console.WriteLine($"VerifyPassword(\"\"): {protection.VerifyPassword("")}");
                Console.WriteLine($"VerifyPassword(null): {protection.VerifyPassword(null)}");

                // Save the workbook to a file
                string filePath = "WorksheetEmptyPasswordProtection.xlsx";
                workbook.Save(filePath);

                // Verify file exists before loading
                if (File.Exists(filePath))
                {
                    // Load the saved workbook to verify persisted protection state
                    Workbook loadedWorkbook = new Workbook(filePath);
                    Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                    Protection loadedProtection = loadedSheet.Protection;

                    // Observe protection status after loading
                    Console.WriteLine("\nAfter loading saved workbook:");
                    Console.WriteLine($"Worksheet.IsProtected: {loadedSheet.IsProtected}");
                    Console.WriteLine($"Protection.IsProtectedWithPassword: {loadedProtection.IsProtectedWithPassword}");
                    Console.WriteLine($"VerifyPassword(\"\"): {loadedProtection.VerifyPassword("")}");
                    Console.WriteLine($"VerifyPassword(null): {loadedProtection.VerifyPassword(null)}");
                }
                else
                {
                    Console.WriteLine($"Error: File '{filePath}' was not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetEmptyPasswordProtectionDemo.Run();
        }
    }
}
