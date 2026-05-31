using System;
using System.IO;
using Aspose.Cells;

namespace ChangeWorkbookPassword
{
    class Program
    {
        static void Main()
        {
            // Path to the existing encrypted workbook
            string inputPath = "EncryptedWorkbook.xlsx";

            // Old (current) password and the new stronger password
            string oldPassword = "oldPass123";
            string newPassword = "NewStrongPassword!@#2026";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: File not found – {inputPath}");
                    return;
                }

                // Load the workbook using the old password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = oldPassword
                };
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Set the new encryption password
                workbook.Settings.Password = newPassword;

                // Save the workbook (overwrites the original file)
                workbook.Save(inputPath);

                Console.WriteLine("Password changed successfully.");
            }
            catch (CellsException ex)
            {
                // Handles Aspose.Cells specific errors (e.g., invalid password)
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handles any other unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}