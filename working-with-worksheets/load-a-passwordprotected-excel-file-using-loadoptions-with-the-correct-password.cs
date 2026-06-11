using System;
using Aspose.Cells;

class LoadPasswordProtectedWorkbook
{
    static void Main()
    {
        // Path to the password‑protected Excel file
        string filePath = "protected.xlsx";

        // Create LoadOptions and set the password required to open the file
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "test";

        // Load the workbook using the LoadOptions with the correct password
        Workbook workbook = new Workbook(filePath, loadOptions);

        // (Optional) Remove the password protection after loading
        workbook.Settings.Password = null;

        // Save the workbook without password protection
        workbook.Save("unprotected.xlsx");
    }
}