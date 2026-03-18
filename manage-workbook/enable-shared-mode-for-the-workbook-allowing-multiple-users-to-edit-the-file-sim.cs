using System;
using Aspose.Cells;

class EnableSharedWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Enable shared mode so multiple users can edit simultaneously
        workbook.Settings.Shared = true;

        // Save the workbook to a file
        string outputPath = "SharedWorkbook.xlsx";
        workbook.Save(outputPath);

        // Load the saved workbook to verify the Shared property
        Workbook loadedWorkbook = new Workbook(outputPath);
        Console.WriteLine("Shared mode enabled: " + loadedWorkbook.Settings.Shared);
    }
}