using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        // Checks if the password satisfies basic complexity requirements:
        // - Minimum 8 characters
        // - At least one uppercase letter
        // - At least one lowercase letter
        // - At least one digit
        // - At least one special character
        static bool IsPasswordComplex(string password)
        {
            if (string.IsNullOrEmpty(password) || password.Length < 8)
                return false;

            bool hasUpper = false, hasLower = false, hasDigit = false, hasSpecial = false;

            foreach (char c in password)
            {
                if (char.IsUpper(c)) hasUpper = true;
                else if (char.IsLower(c)) hasLower = true;
                else if (char.IsDigit(c)) hasDigit = true;
                else if (!char.IsWhiteSpace(c) && !char.IsLetterOrDigit(c)) hasSpecial = true;
            }

            return hasUpper && hasLower && hasDigit && hasSpecial;
        }

        static void Main(string[] args)
        {
            // Define a strong password
            string password = "Str0ngP@ssw0rd!";

            // Verify password complexity before applying it
            if (!IsPasswordComplex(password))
            {
                Console.WriteLine("Password does not meet complexity requirements.");
                return;
            }

            // Create a new workbook (lifecycle rule: create)
            Workbook wb = new Workbook();

            // Add sample data
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encrypted workbook example");

            // Apply password protection (lifecycle rule: set property)
            wb.Settings.Password = password;

            // Set strong encryption options (lifecycle rule: method)
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the encrypted workbook (lifecycle rule: save)
            string filePath = "EncryptedWorkbook.xlsx";
            wb.Save(filePath, SaveFormat.Xlsx);

            // Verify that the workbook is encrypted
            Console.WriteLine($"Workbook encrypted: {wb.Settings.IsEncrypted}");

            // Load the workbook using the password (lifecycle rule: load)
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = password;
            Workbook loadedWb = new Workbook(filePath, loadOptions);

            // Confirm that the loaded workbook is encrypted and accessible
            Console.WriteLine($"Loaded workbook encrypted: {loadedWb.Settings.IsEncrypted}");
            Console.WriteLine($"Cell A1 value after loading: {loadedWb.Worksheets[0].Cells["A1"].StringValue}");

            // Additional verification using FileFormatUtil
            using (Stream stream = File.OpenRead(filePath))
            {
                bool isPasswordValid = FileFormatUtil.VerifyPassword(stream, password);
                Console.WriteLine($"Password validation via FileFormatUtil: {isPasswordValid}");
            }
        }
    }
}