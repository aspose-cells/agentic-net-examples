// Title: C# Aspose.Cells example: catch exception when unprotecting a worksheet with a wrong password
// AI Prompts: Generate C# code that protects an Aspose.Cells worksheet, attempts to unprotect it with an incorrect password, and captures the thrown exception. | Show how to implement try‑catch around Worksheet.Unprotect to handle invalid password errors in Aspose.Cells for .NET. | Demonstrate saving a workbook after a failed unprotect attempt while preserving exception details in C#.
// Common Searches: how to catch Aspose.Cells unprotect worksheet wrong password exception c# | Aspose.Cells worksheet.Unprotect throws error with invalid password example | C# code to protect Excel sheet and handle wrong password using Aspose.Cells | saving workbook after failed unprotect attempt Aspose.Cells .NET | exception handling for worksheet protection Aspose.Cells C# tutorial
// Tags: Aspose.Cells worksheet unprotect exception handling | C# protect worksheet with password Aspose.Cells | invalid password error Aspose.Cells worksheet | save workbook after protection Aspose.Cells C# | try-catch Worksheet.Unprotect Aspose.Cells | Aspose.Cells encryption protection example C#

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a workbook, protects the first worksheet with a password, then tries to unprotect it using an incorrect password. The resulting exception is caught and displayed, and the workbook is saved afterward.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "ProtectedSheet";

            // Protect the worksheet with a known password.
            // The third parameter (oldPassword) is not required here, so pass null.
            worksheet.Protect(ProtectionType.All, "correctPassword", null);

            try
            {
                // Attempt to unprotect the worksheet using an incorrect password
                worksheet.Unprotect("wrongPassword");
                Console.WriteLine("Worksheet unprotected (unexpected).");
            }
            catch (Exception ex)
            {
                // Capture and display the exception thrown due to wrong password
                Console.WriteLine("Exception caught while unprotecting worksheet:");
                Console.WriteLine(ex.Message);
            }

            // Optional: Save the workbook if needed
            string outputPath = "ProtectedDemo.xlsx";
            try
            {
                // Ensure the directory exists before saving
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Error saving workbook:");
                Console.WriteLine(saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected errors
            Console.WriteLine("An unexpected error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
