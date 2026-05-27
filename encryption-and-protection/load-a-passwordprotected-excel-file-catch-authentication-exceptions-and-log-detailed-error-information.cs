using System;
using Aspose.Cells;

namespace AsposeCellsPasswordDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the password‑protected Excel file
            string filePath = "protected.xlsx";

            // The password to try opening the file with
            string password = "wrongPassword";

            // Configure load options with the supplied password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = password;

            try
            {
                // Attempt to load the workbook using the load options
                Workbook workbook = new Workbook(filePath, loadOptions);

                // If loading succeeds, read a sample cell to verify access
                object cellValue = workbook.Worksheets[0].Cells["A1"].Value;
                Console.WriteLine($"Cell A1 value: {cellValue}");
            }
            // Catch authentication failures specifically
            catch (CellsException ex) when (ex.Code == ExceptionType.IncorrectPassword)
            {
                Console.WriteLine("Authentication failed: Incorrect password.");
                Console.WriteLine($"Exception Code: {ex.Code}");
                Console.WriteLine($"Message: {ex.Message}");
                Console.WriteLine($"StackTrace: {ex.StackTrace}");
            }
            // Catch any other unexpected exceptions
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred while loading the workbook.");
                Console.WriteLine($"Message: {ex.Message}");
                Console.WriteLine($"StackTrace: {ex.StackTrace}");
            }
        }
    }
}