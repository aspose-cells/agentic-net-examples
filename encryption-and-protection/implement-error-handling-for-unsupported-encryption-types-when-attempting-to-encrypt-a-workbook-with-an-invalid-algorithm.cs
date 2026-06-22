using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set a password for encryption
            workbook.Settings.Password = "myPassword";

            // Define an invalid encryption type (value not defined in the enum)
            EncryptionType invalidEncryptionType = (EncryptionType)999;

            try
            {
                // Attempt to apply encryption options with the invalid type
                workbook.SetEncryptionOptions(invalidEncryptionType, 128);

                // Save the workbook (this line will not be reached if an exception is thrown above)
                workbook.Save("EncryptedWorkbook.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.UnsupportedFeature)
            {
                // Handle the specific case where the encryption algorithm is not supported
                Console.WriteLine("Error: The specified encryption type is not supported.");
            }
            catch (Exception ex)
            {
                // General fallback for any other unexpected errors
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}