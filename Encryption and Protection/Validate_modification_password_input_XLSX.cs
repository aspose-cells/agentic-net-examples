using System;
using Aspose.Cells;

namespace AsposeCellsPasswordValidation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input XLSX file
            string filePath = "input.xlsx";

            // Password to validate (could be obtained from user input)
            string passwordToValidate = "owner";

            // Load the workbook (no password needed for write‑protected files)
            Workbook workbook = new Workbook(filePath);

            // Validate the write‑protection (modification) password
            bool isValid = workbook.Settings.WriteProtection.ValidatePassword(passwordToValidate);

            // Output the validation result
            Console.WriteLine($"Password '{passwordToValidate}' is valid: {isValid}");
        }
    }
}