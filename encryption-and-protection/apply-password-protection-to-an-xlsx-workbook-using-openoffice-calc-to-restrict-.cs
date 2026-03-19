using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add some data to the first worksheet
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive data");

        // Set a password to encrypt the workbook (prevents opening without password)
        wb.Settings.Password = "OpenCalcPass";

        // Save the password‑protected workbook as XLSX
        wb.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);

        // Load the protected workbook using the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "OpenCalcPass";
        Workbook wbProtected = new Workbook("ProtectedWorkbook.xlsx", loadOptions);

        // Verify that the data can be accessed after providing the correct password
        Console.WriteLine("Cell A1 value: " + wbProtected.Worksheets[0].Cells["A1"].StringValue);
    }
}