using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    class OpenEncryptedWorkbook
    {
        static void Main()
        {
            // Path to the encrypted Excel file
            string filePath = "encrypted.xlsx";

            // Password required to open the workbook
            string password = "myPassword";

            // Verify that the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Create LoadOptions with the appropriate format and password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };

                // Load the password‑protected workbook using the LoadOptions
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Example: read a cell value to confirm successful loading
                Console.WriteLine("Cell A1 value: " + workbook.Worksheets[0].Cells["A1"].Value);
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