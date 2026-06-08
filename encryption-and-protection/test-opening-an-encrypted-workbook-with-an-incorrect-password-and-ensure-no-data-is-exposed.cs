using System;
using Aspose.Cells;

class TestEncryptedWorkbook
{
    static void Main()
    {
        // Create a new workbook and add some confidential data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Secret Data");

        // Protect the workbook with a password
        wb.Settings.Password = "correctPassword";

        // Save the encrypted workbook to disk
        string filePath = "encryptedWorkbook.xlsx";
        wb.Save(filePath);

        // Attempt to open the encrypted workbook with an incorrect password
        try
        {
            LoadOptions wrongOptions = new LoadOptions();
            wrongOptions.Password = "wrongPassword";

            // This should throw an exception because the password is invalid
            Workbook wbWrong = new Workbook(filePath, wrongOptions);

            // If no exception occurs, data exposure has happened (should not happen)
            string exposedValue = wbWrong.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Unexpectedly opened workbook. Cell value: " + exposedValue);
        }
        catch (Exception ex)
        {
            // Expected outcome: the workbook cannot be opened with the wrong password
            Console.WriteLine("Failed to open workbook with incorrect password: " + ex.Message);
        }

        // Verify that the workbook can be opened with the correct password
        LoadOptions correctOptions = new LoadOptions();
        correctOptions.Password = "correctPassword";

        Workbook wbCorrect = new Workbook(filePath, correctOptions);
        Console.WriteLine("Opened with correct password. Cell value: " + wbCorrect.Worksheets[0].Cells["A1"].StringValue);
    }
}