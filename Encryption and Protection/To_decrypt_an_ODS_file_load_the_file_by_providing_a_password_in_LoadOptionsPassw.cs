using System;
using Aspose.Cells;

class DecryptOdsExample
{
    static void Main()
    {
        Run();
    }

    public static void Run()
    {
        const string password = "your_password";

        // Create a sample workbook and protect it as ODS
        var wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Sample data");
        wb.Settings.Password = password; // protect workbook
        const string protectedFile = "protected.ods";
        wb.Save(protectedFile, SaveFormat.Ods);

        // Load the encrypted ODS workbook with the specified password
        var loadOptions = new OdsLoadOptions
        {
            Password = password
        };
        var protectedWorkbook = new Workbook(protectedFile, loadOptions);

        // Save the workbook as an unprotected XLSX file
        protectedWorkbook.Save("decrypted.xlsx", SaveFormat.Xlsx);
    }
}