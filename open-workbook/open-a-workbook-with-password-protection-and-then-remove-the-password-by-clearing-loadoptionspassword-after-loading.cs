using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the password‑protected workbook using LoadOptions
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "test"; // password of the source file

        Workbook workbook = new Workbook("protected.xlsx", loadOptions);

        // Remove the workbook password by clearing the Settings.Password property
        workbook.Settings.Password = null;

        // Save the workbook without password protection
        workbook.Save("unprotected.xlsx");
    }
}

// Author note: This example demonstrates opening a protected workbook and removing its password.