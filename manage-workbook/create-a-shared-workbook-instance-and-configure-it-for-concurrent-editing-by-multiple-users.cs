using System;
using Aspose.Cells;

class SharedWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Enable shared mode to allow concurrent editing by multiple users
        workbook.Settings.Shared = true;

        // (Optional) Protect the shared workbook with a password
        // workbook.ProtectSharedWorkbook("myPassword");

        // Save the shared workbook to disk
        string outputPath = "SharedWorkbook.xlsx";
        workbook.Save(outputPath);

        // Load the workbook back to verify that the Shared property is set
        Workbook loadedWorkbook = new Workbook(outputPath);
        Console.WriteLine("Shared property value: " + loadedWorkbook.Settings.Shared);
    }
}