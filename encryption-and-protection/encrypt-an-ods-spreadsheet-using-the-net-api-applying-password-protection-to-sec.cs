using System;
using Aspose.Cells;

class EncryptOdsExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Access the first worksheet and add some data
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive data");

        // Set password to protect (encrypt) the workbook
        wb.Settings.Password = "mySecretPassword";

        // Save the workbook as an ODS file with password protection
        wb.Save("encrypted_output.ods");

        // Load the password‑protected workbook using the same password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "mySecretPassword";
        Workbook loadedWb = new Workbook("encrypted_output.ods", loadOptions);

        // Verify that the data can be read after providing the password
        Console.WriteLine("Loaded cell value: " + loadedWb.Worksheets[0].Cells["A1"].Value);
    }
}