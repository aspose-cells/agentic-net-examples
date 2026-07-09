using System;
using Aspose.Cells;

namespace AsposeCellsPasswordDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the password‑protected Excel file
            string filePath = "protected.xlsx";

            // The password to use for opening the file (intentionally incorrect to trigger the exception)
            string password = "wrongPassword";

            // Configure load options with the supplied password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = password;

            try
            {
                // Attempt to load the workbook using the load options
                Workbook workbook = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook loaded successfully.");
                // Additional processing can be placed here
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.IncorrectPassword)
            {
                // Specific handling for authentication failures
                Console.WriteLine("Authentication failed: Incorrect password.");
                Console.WriteLine($"Exception Code   : {ex.Code}");
                Console.WriteLine($"Exception Message: {ex.Message}");
                Console.WriteLine($"Stack Trace       : {ex.StackTrace}");
            }
            catch (CellsException ex)
            {
                // General Aspose.Cells related exceptions
                Console.WriteLine("A CellsException was caught.");
                Console.WriteLine($"Exception Code   : {ex.Code}");
                Console.WriteLine($"Exception Message: {ex.Message}");
                Console.WriteLine($"Stack Trace       : {ex.StackTrace}");
            }
            catch (Exception ex)
            {
                // Any other unexpected exceptions
                Console.WriteLine("An unexpected exception occurred.");
                Console.WriteLine($"Message   : {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            }
        }
    }
}