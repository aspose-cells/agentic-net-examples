// Title: Create a password‑protected Excel (.xlsx) file with Aspose.Cells for .NET and verify that loading it without a password throws an exception
// AI Prompts: Create a new Workbook, add sample cells, set Workbook.Settings.Password, and save the file as an encrypted .xlsx. | Open the saved .xlsx without supplying a password, catch the thrown exception, and output its message. | Print the full path of the encrypted workbook and confirm the exception handling flow.
// Common Searches: how to set a password on an xlsx workbook using Aspose.Cells in C# | c# attempt to open a password‑protected Excel file without a password and get exception | test Aspose.Cells encrypted Excel file compatibility with other .NET libraries | sample code for creating password protected Excel file with Aspose.Cells for .NET | exception message when loading encrypted workbook without password Aspose.Cells
// Tags: Aspose.Cells workbook password protection | encrypt xlsx file using Aspose.Cells | load encrypted workbook without password exception | password protection compatibility test Aspose.Cells | Aspose.Cells workbook encryption .NET

using System;
using System.IO;
using Aspose.Cells;

// The program builds a new workbook, writes sample text and the current date, applies a password via Workbook.Settings.Password, saves the file as an encrypted .xlsx, then attempts to load the file without providing the password, catches the expected exception, and prints the file path and error details.
class Program
{
    static void Main()
    {
        try
        {
            // Path for the Excel file
            string filePath = "encrypted.xlsx";

            // ------------------------------
            // Create a workbook with Aspose.Cells
            // ------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Add some sample data
            sheet.Cells["A1"].PutValue("Hello, Aspose!");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Encrypt the workbook with a password
            workbook.Settings.Password = "myPassword";

            // Save the encrypted file
            workbook.Save(filePath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved and encrypted at: {Path.GetFullPath(filePath)}");

            // ------------------------------
            // Attempt to open the encrypted file without providing a password
            // ------------------------------
            try
            {
                // Ensure the file exists before attempting to load
                if (!File.Exists(filePath))
                    throw new FileNotFoundException("Encrypted file not found.", filePath);

                // This will throw because the workbook is password‑protected
                Workbook openedWorkbook = new Workbook(filePath);
                Console.WriteLine("Opened encrypted file without password (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to open encrypted file without password as expected.");
                Console.WriteLine($"Exception: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
