using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the workbook author metadata
        workbook.Settings.Author = "John Doe";

        // Apply write protection and set its author
        workbook.Settings.WriteProtection.Author = "John Doe";
        workbook.Settings.WriteProtection.Password = "protect123";
        workbook.Settings.WriteProtection.RecommendReadOnly = true;

        // Save the workbook with write protection
        string filePath = "ProtectedWorkbook.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);

        // Load the workbook to verify the settings
        Workbook loaded = new Workbook(filePath);
        Console.WriteLine("Workbook Author: " + loaded.Settings.Author);
        Console.WriteLine("Write Protection Author: " + loaded.Settings.WriteProtection.Author);
        Console.WriteLine("Is Write Protected: " + loaded.Settings.WriteProtection.IsWriteProtected);
    }
}