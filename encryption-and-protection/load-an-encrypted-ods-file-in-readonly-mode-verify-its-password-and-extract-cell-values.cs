using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the encrypted ODS file
        string filePath = "encrypted.ods";

        // Password to open the file
        string password = "myPassword";

        // Verify the password using a file stream
        bool isPasswordCorrect;
        using (FileStream stream = File.OpenRead(filePath))
        {
            // Returns true if the password matches the encrypted file
            isPasswordCorrect = FileFormatUtil.VerifyPassword(stream, password);
        }

        Console.WriteLine($"Password verification result: {isPasswordCorrect}");

        if (!isPasswordCorrect)
        {
            Console.WriteLine("Invalid password. Unable to load the workbook.");
            return;
        }

        // Load the ODS workbook with the verified password
        OdsLoadOptions loadOptions = new OdsLoadOptions();
        loadOptions.Password = password; // Set password in load options

        // Open the workbook (read‑only usage – we do not modify it)
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Extract and display cell values from all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet: {sheet.Name}");
            Cells cells = sheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var cell = cells[row, col];
                    if (cell.Value != null)
                    {
                        Console.WriteLine($"Cell {cell.Name}: {cell.Value}");
                    }
                }
            }
        }
    }
}