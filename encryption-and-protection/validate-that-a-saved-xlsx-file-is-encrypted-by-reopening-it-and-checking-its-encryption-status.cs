using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path and password for the encrypted workbook
        string filePath = "encrypted.xlsx";
        string password = "mySecret";

        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Encrypted test");

        // Encrypt the workbook with a password
        workbook.Settings.Password = password;

        // Save the encrypted workbook
        workbook.Save(filePath);

        // Reopen the file and check its encryption status using FileFormatInfo
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
        Console.WriteLine("Is the saved file encrypted? " + formatInfo.IsEncrypted);

        // Load the encrypted workbook with the correct password to verify it can be opened
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
        loadOptions.Password = password;
        Workbook loadedWorkbook = new Workbook(filePath, loadOptions);
        Console.WriteLine("Loaded workbook cell A1 value: " + loadedWorkbook.Worksheets[0].Cells["A1"].StringValue);
    }
}