// Title: Protect an Aspose.Cells worksheet with an empty password and verify after save
// Description: C# example that creates a workbook, applies Worksheet.Protect with an empty string, displays IsProtected, Protection.IsProtectedWithPassword, and password length, verifies the empty password, saves the file, reloads it, and repeats the checks to show how Aspose.Cells records protection when no password is supplied.
// Keywords: Aspose.Cells worksheet protection | C# Protect empty password | Worksheet.IsProtected | Protection.IsProtectedWithPassword | VerifyPassword empty string | save protected workbook Aspose.Cells | reload workbook protection status | Aspose.Cells encryption and protection
// Common Searches: Aspose.Cells protect worksheet with no password | What does Worksheet.IsProtected return when password is empty? | How to verify an empty password on a protected sheet in Aspose.Cells | Does saving a workbook keep protection if password is blank? | C# example for empty‑password worksheet protection Aspose.Cells
// Developer Intent: Show how to protect a worksheet using an empty password, inspect protection flags, verify the password, and confirm that the protection persists after saving and reloading the workbook.
// Use Cases: Apply worksheet protection without requiring a user‑entered password while still preventing edits through the UI. | Understand the difference between IsProtected and Protection.IsProtectedWithPassword when the password is omitted. | Validate that protection settings survive the save/load cycle for automated processing pipelines.
// AI Prompts: Write C# code that uses Aspose.Cells to protect a worksheet with an empty password, prints protection properties, saves the workbook, reloads it, and re‑prints the properties. | Explain the expected values of Worksheet.IsProtected, Protection.IsProtectedWithPassword, and password length when Protect is called with "". | Provide steps to verify an empty password on a worksheet after loading a workbook that was saved with empty‑password protection.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, applies Worksheet.Protect with an empty string, displays IsProtected, Protection.IsProtectedWithPassword, and password length, verifies the empty password, saves the file, reloads it, and repeats the checks to show how Aspose.Cells records protection when no password is supplied.
    public class EmptyPasswordProtectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Protect the worksheet with an empty password string
                sheet.Protect(ProtectionType.All, "", null);

                // Display protection information
                Console.WriteLine("Worksheet.IsProtected: " + sheet.IsProtected);
                Console.WriteLine("Worksheet.Protection.IsProtectedWithPassword: " + sheet.Protection.IsProtectedWithPassword);
                Console.WriteLine("Worksheet.Protection.Password length: " + (sheet.Protection.Password?.Length ?? 0));

                // Verify the empty password
                bool verifyEmpty = sheet.Protection.VerifyPassword("");
                Console.WriteLine("Verify empty password: " + verifyEmpty);

                // Save the workbook
                string filePath = "EmptyPasswordProtected.xlsx";
                workbook.Save(filePath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(filePath)}");

                // Load the saved workbook and check protection status again
                if (File.Exists(filePath))
                {
                    Workbook loadedWorkbook = new Workbook(filePath);
                    Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

                    Console.WriteLine("Loaded Worksheet.IsProtected: " + loadedSheet.IsProtected);
                    Console.WriteLine("Loaded Worksheet.Protection.IsProtectedWithPassword: " + loadedSheet.Protection.IsProtectedWithPassword);
                    bool verifyEmptyLoaded = loadedSheet.Protection.VerifyPassword("");
                    Console.WriteLine("Loaded Verify empty password: " + verifyEmptyLoaded);
                }
                else
                {
                    Console.WriteLine($"File not found: {filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            EmptyPasswordProtectionDemo.Run();
        }
    }
}
