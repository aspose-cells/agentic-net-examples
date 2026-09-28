// Title: Load an Excel workbook with Aspose.Cells for .NET and log each worksheet’s paper size setting
// AI Prompts: Write C# code that opens a given .xlsx file using Aspose.Cells, iterates through all worksheets, and prints each worksheet’s name together with its PageSetup.PaperSize value. | Show how to add error handling for missing files and workbook‑load failures while retrieving and logging the paper size of every worksheet in a console application.
// Common Searches: Aspose.Cells C# get worksheet paper size setting from an Excel workbook | Enumerate worksheets and display their paper size values using Aspose.Cells .NET | Console application to list each sheet’s name and paper size with Aspose.Cells | Detect automatic paper size for worksheets when loading an .xlsx file with Aspose.Cells
// Tags: Aspose.Cells load workbook and read worksheet page setup | C# retrieve worksheet PaperSize using Aspose.Cells | enumerate worksheets and log paper size Aspose.Cells | handle missing Excel file Aspose.Cells .NET | error handling workbook loading Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program checks that 'input.xlsx' exists, loads it with Aspose.Cells, iterates through every worksheet, reads the PageSetup.PaperSize property, and writes each worksheet name and its paper size to the console, with robust error handling for missing files and load failures.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the Excel workbook from the specified file
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Retrieve the paper size setting for the worksheet
                PaperSizeType paperSize = sheet.PageSetup.PaperSize;

                // Log the worksheet name and its paper size
                Console.WriteLine($"Worksheet '{sheet.Name}': Paper Size = {paperSize}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}
