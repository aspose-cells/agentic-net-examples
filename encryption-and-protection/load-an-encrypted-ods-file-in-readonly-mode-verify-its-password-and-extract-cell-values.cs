using System;
using System.IO;
using Aspose.Cells;

class LoadEncryptedOds
{
    static void Main()
    {
        // Path to the encrypted ODS file and its password
        string filePath = "protected.ods";
        string password = "secret";

        // Verify the password before attempting to load the file
        bool passwordValid;
        using (FileStream stream = File.OpenRead(filePath))
        {
            passwordValid = FileFormatUtil.VerifyPassword(stream, password);
        }

        Console.WriteLine($"Password verification result: {passwordValid}");
        if (!passwordValid)
        {
            Console.WriteLine("Invalid password. Exiting.");
            return;
        }

        // Load the ODS file with the verified password using OdsLoadOptions
        OdsLoadOptions loadOptions = new OdsLoadOptions();
        loadOptions.Password = password; // Set password for loading

        Workbook workbook = new Workbook(filePath, loadOptions);

        // Confirm that the workbook reports being encrypted
        Console.WriteLine($"Workbook IsEncrypted: {workbook.Settings.IsEncrypted}");

        // Extract and display cell values from the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                var value = cells[row, col].Value;
                if (value != null)
                {
                    Console.WriteLine($"Cell {cells[row, col].Name}: {value}");
                }
            }
        }
    }
}