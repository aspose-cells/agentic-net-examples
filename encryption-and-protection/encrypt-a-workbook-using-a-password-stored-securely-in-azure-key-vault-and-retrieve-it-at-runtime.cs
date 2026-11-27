// Title: Encrypt an Excel workbook at runtime using a password from an environment variable with Aspose.Cells for .NET
// Common Searches: how to set a password on an Aspose.Cells workbook in C# before saving | retrieve workbook encryption password from environment variable in .NET | save password-protected .xlsx file using Aspose.Cells C# example | protect Excel file structure and windows programmatically with Aspose.Cells | fallback default password for Aspose.Cells workbook protection in C#
// Tags: Aspose.Cells workbook password protection C# | encrypt Excel file at runtime Aspose.Cells | environment variable password retrieval .NET | save password-protected .xlsx using Aspose.Cells | protect workbook structure Aspose.Cells C#

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionExample
{
    // The sample reads a password from the WORKBOOK_PASSWORD environment variable (or uses a default), creates a Workbook, adds sample data, assigns the password to workbook.Settings.Password to require it on opening, and saves the file as an encrypted EncryptedWorkbook.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Retrieve the password from an environment variable.
                // If not set, fall back to a default password.
                string workbookPassword = Environment.GetEnvironmentVariable("WORKBOOK_PASSWORD");
                if (string.IsNullOrEmpty(workbookPassword))
                {
                    workbookPassword = "DefaultPassword123";
                }

                // Create a new workbook.
                var workbook = new Workbook();

                // Add some sample data to the first worksheet.
                var worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("Hello");
                worksheet.Cells["B1"].PutValue("World");
                worksheet.Cells["A2"].PutValue(DateTime.Now);
                worksheet.Cells["B2"].PutValue(12345);

                // Protect the workbook with the retrieved password.
                // This will require the password to open the file.
                workbook.Settings.Password = workbookPassword;

                // Optionally, you can also protect the structure and windows.
                // workbook.Protect(ProtectionType.All, workbookPassword);

                // Save the encrypted workbook.
                string outputPath = "EncryptedWorkbook.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved and encrypted at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
