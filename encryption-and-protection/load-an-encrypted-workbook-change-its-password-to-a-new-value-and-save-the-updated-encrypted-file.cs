// Title: Load an encrypted Excel workbook, change its password, and save it with a new password using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells change password of an existing encrypted .xlsx file | Load password‑protected Excel workbook with Aspose.Cells and set a new password | Re‑encrypt an Excel file with a different password using Aspose.Cells .NET | Aspose.Cells LoadOptions Password property to open encrypted workbook in C# | How to update workbook.Settings.Password and save encrypted Excel in .NET
// Tags: Aspose.Cells load encrypted workbook C# | Aspose.Cells change workbook password | Aspose.Cells re‑encrypt Excel file .NET | LoadOptions Password property Aspose.Cells | Workbook.Settings.Password usage

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a password‑protected workbook if none exists, opens it with LoadOptions using the current password, updates Workbook.Settings.Password to a new value, and saves the file, demonstrating how to re‑encrypt an Excel workbook with a different password in C# while handling common exceptions.
class Program
{
    static void Main()
    {
        // Input encrypted workbook file
        string inputFile = "encrypted_input.xlsx";

        // Current password of the workbook
        string oldPassword = "oldPassword123";

        // Desired new password
        string newPassword = "newPassword456";

        // Output file path for the re‑encrypted workbook
        string outputFile = "encrypted_output.xlsx";

        try
        {
            // If the input file does not exist, create a sample workbook and encrypt it with the old password
            if (!File.Exists(inputFile))
            {
                var sampleWorkbook = new Workbook();
                var sheet = sampleWorkbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample data");
                // Set password for saving
                sampleWorkbook.Settings.Password = oldPassword;
                sampleWorkbook.Save(inputFile, SaveFormat.Xlsx);
                Console.WriteLine($"Sample encrypted workbook created: {inputFile}");
            }

            // Load the workbook using the existing password
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = oldPassword // set password for opening
            };
            var workbook = new Workbook(inputFile, loadOptions);

            // Change the password to the new value
            workbook.Settings.Password = newPassword; // set password for saving

            // Save the workbook; it will be encrypted with the new password
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook re‑encrypted and saved as: {outputFile}");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine($"File not found: {fnfEx.FileName}");
        }
        catch (CellsException cellsEx)
        {
            Console.WriteLine($"Aspose.Cells error: {cellsEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
