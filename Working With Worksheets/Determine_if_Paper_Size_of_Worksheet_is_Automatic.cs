using System;
using Aspose.Cells;

class DetermineAutomaticPaperSize
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the PageSetup object
        PageSetup pageSetup = worksheet.PageSetup;

        // Check if the paper size is set to automatic
        bool isAutomatic = pageSetup.IsAutomaticPaperSize;
        Console.WriteLine("Automatic Paper Size (new workbook): " + isAutomatic);

        // Attempt to load an existing workbook and check its setting
        // Replace "input.xlsx" with a valid file path if available
        try
        {
            Workbook loadedWorkbook = new Workbook("input.xlsx");
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
            bool isAutomaticLoaded = loadedWorksheet.PageSetup.IsAutomaticPaperSize;
            Console.WriteLine("Automatic Paper Size (loaded workbook): " + isAutomaticLoaded);
        }
        catch (System.IO.FileNotFoundException)
        {
            Console.WriteLine("input.xlsx not found; load example skipped.");
        }

        // Save the new workbook (optional)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}