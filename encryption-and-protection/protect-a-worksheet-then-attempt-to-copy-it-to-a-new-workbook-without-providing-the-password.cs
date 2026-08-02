// Title: Aspose.Cells for .NET – Attempt to Copy a Password‑Protected Worksheet Without Supplying the Password (C#)
// Description: This C# example creates a workbook, protects its first worksheet with a password using Worksheet.Protect, then tries to copy that sheet to a new workbook with Worksheets.AddCopy while omitting the password. The code catches the resulting exception, logs the error, and saves both workbooks for inspection.
// Keywords: Aspose.Cells | C# | worksheet protection | Protect method | AddCopy | copy protected sheet | password‑protected worksheet | exception handling | copy without password | .NET example
// Common Searches: Aspose.Cells copy protected worksheet without password | Worksheets.AddCopy exception when sheet is protected | How to copy a password‑protected sheet using Aspose.Cells .NET | Why does AddCopy fail on a protected worksheet | C# Aspose.Cells worksheet protection copy error
// Developer Intent: Show that copying a password‑protected worksheet without providing the password triggers an exception, demonstrating the need to supply the password or remove protection first.
// Use Cases: Verify that Worksheets.AddCopy throws an error when the source sheet is protected and no password is given. | Implement robust error handling and logging for failed copy operations on protected worksheets. | Save the source and destination workbooks to examine protection status after an attempted copy.
// AI Prompts: Generate C# code with Aspose.Cells that copies a protected worksheet to another workbook by providing the required password. | Explain why Worksheets.AddCopy cannot duplicate a password‑protected sheet without the password and outline the correct steps to achieve the copy. | Create a C# unit test that asserts an exception is thrown when AddCopy is called on a protected worksheet without supplying a password.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWorksheetProtectionDemo
{
    // This C# example creates a workbook, protects its first worksheet with a password using Worksheet.Protect, then tries to copy that sheet to a new workbook with Worksheets.AddCopy while omitting the password. The code catches the resulting exception, logs the error, and saves both workbooks for inspection.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create source workbook and protect its first worksheet with a password
                Workbook sourceWorkbook = new Workbook();
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
                sourceSheet.Protect(ProtectionType.All, "SecretPwd", null);
                Console.WriteLine($"Source worksheet protected: {sourceSheet.IsProtected}");

                // Attempt to copy the protected worksheet to a new workbook without providing the password
                Workbook destWorkbook = new Workbook();
                try
                {
                    // AddCopy expects the worksheet name, not the Worksheet object
                    destWorkbook.Worksheets.AddCopy(sourceSheet.Name);
                    Console.WriteLine("Worksheet copied successfully (unexpected).");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to copy protected worksheet without password: {ex.Message}");
                }

                // Save the workbooks for inspection (optional)
                SaveWorkbook(sourceWorkbook, "ProtectedSource.xlsx");
                SaveWorkbook(destWorkbook, "Destination.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Helper method to save a workbook safely
        private static void SaveWorkbook(Workbook workbook, string filePath)
        {
            try
            {
                // Ensure the directory exists
                string directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(filePath);
                Console.WriteLine($"Workbook saved: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook '{filePath}': {ex.Message}");
            }
        }
    }
}
