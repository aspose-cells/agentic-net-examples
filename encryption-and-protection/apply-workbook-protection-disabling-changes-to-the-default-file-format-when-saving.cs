using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable write protection to prevent changes to the default file format when saving
        workbook.Settings.WriteProtection.Password = "mySecretPwd";
        workbook.Settings.WriteProtection.RecommendReadOnly = true;

        // Save the workbook in the default format (XLSX)
        workbook.Save("ProtectedWorkbook.xlsx");

        // Load the saved workbook to verify the protection settings
        Workbook loadedWorkbook = new Workbook("ProtectedWorkbook.xlsx");
        bool isWriteProtected = loadedWorkbook.Settings.WriteProtection.IsWriteProtected;
        Console.WriteLine("Workbook is write protected: " + isWriteProtected);
    }
}