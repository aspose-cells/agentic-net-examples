using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and put some data in the first cell
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Secret data");

        // Set a password to encrypt the workbook
        wb.Settings.Password = "mySecret";

        // Save the encrypted workbook to disk
        string filePath = "encrypted.xlsx";
        wb.Save(filePath);

        // Attempt to open the encrypted workbook without providing a password
        // This should throw an exception because the file is encrypted
        try
        {
            Workbook wbWithoutPassword = new Workbook(filePath);
            Console.WriteLine("Opened without password (unexpected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to open without password as expected: " + ex.Message);
        }

        // Open the encrypted workbook with the correct password using LoadOptions
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "mySecret";
        Workbook wbWithPassword = new Workbook(filePath, loadOptions);

        // Verify that the data can be read after providing the correct password
        Console.WriteLine("Opened with password, cell A1 value: " + wbWithPassword.Worksheets[0].Cells["A1"].StringValue);
    }
}