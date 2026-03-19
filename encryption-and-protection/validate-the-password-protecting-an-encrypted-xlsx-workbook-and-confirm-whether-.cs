using System;
using System.IO;
using Aspose.Cells;

class ValidateEncryptedWorkbook
{
    static void Main()
    {
        // Path to the encrypted workbook and the password to test
        string filePath = "encrypted.xlsx";
        string password = "test";

        // Verify the password without fully loading the workbook
        bool isPasswordValid;
        using (Stream stream = File.OpenRead(filePath))
        {
            isPasswordValid = FileFormatUtil.VerifyPassword(stream, password);
        }

        Console.WriteLine($"Password validation result: {isPasswordValid}");

        if (isPasswordValid)
        {
            // Load the workbook using the correct password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = password;
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Example processing: read the value of cell A1
            string cellValue = workbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Cell A1 value: {cellValue}");
        }
        else
        {
            // Access is not authorized
            Console.WriteLine("Access denied: incorrect password.");
        }
    }
}