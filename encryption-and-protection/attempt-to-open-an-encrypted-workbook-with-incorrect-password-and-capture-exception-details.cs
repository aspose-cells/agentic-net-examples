using System;
using Aspose.Cells;

namespace AsposeCellsPasswordDemo
{
    class Program
    {
        static void Main()
        {
            // Path for the encrypted workbook
            string filePath = "encrypted.xlsx";

            // -----------------------------------------------------------------
            // Create a new workbook and add sample data
            // -----------------------------------------------------------------
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // Protect the workbook with a password
            wb.Settings.Password = "correctPassword";

            // Save the encrypted workbook
            wb.Save(filePath);
            Console.WriteLine($"Workbook saved with encryption to '{filePath}'.");

            // -----------------------------------------------------------------
            // Attempt to open the encrypted workbook with an incorrect password
            // -----------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "wrongPassword";

            try
            {
                // This line should throw an exception because the password is incorrect
                Workbook protectedWb = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook opened successfully (unexpected).");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.IncorrectPassword)
            {
                // Capture details specific to incorrect password
                Console.WriteLine("Failed to open workbook: Incorrect password.");
                Console.WriteLine($"Exception Message: {ex.Message}");
                Console.WriteLine($"Exception Code: {ex.Code}");
            }
            catch (Exception ex)
            {
                // Capture any other unexpected exceptions
                Console.WriteLine("An unexpected error occurred while opening the workbook.");
                Console.WriteLine($"Exception Message: {ex.Message}");
            }
        }
    }
}