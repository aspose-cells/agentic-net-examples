// Title: Load an Excel workbook and log each worksheet’s Fit‑to‑Pages (automatic paper size) setting using Aspose.Cells for .NET
// Common Searches: how to check FitToPagesWide and FitToPagesTall for each sheet with Aspose.Cells C# | C# read worksheet page setup to determine automatic paper size using Aspose.Cells | log fit‑to‑pages status of all worksheets in an Excel file with Aspose.Cells | Aspose.Cells detect if worksheets are set to fit to pages in .NET | verify worksheet fit‑to‑pages flag before printing with Aspose.Cells C#
// Tags: Aspose.Cells workbook iteration C# | Aspose.Cells page setup FitToPages detection | C# determine worksheet automatic paper size | Aspose.Cells log fit-to-pages flag | Aspose.Cells file existence check

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, iterates through each Worksheet, checks PageSetup.FitToPagesWide and FitToPagesTall to determine if the sheet is set to fit to pages (automatic paper size), and writes the worksheet name with a boolean status to the console while handling missing‑file and other exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine if the worksheet is set to "Fit to pages"
                // In Aspose.Cells, FitToPages is represented by non‑zero values for both FitToPagesWide and FitToPagesTall
                bool isFitToPages = sheet.PageSetup.FitToPagesWide > 0 && sheet.PageSetup.FitToPagesTall > 0;

                // Log the worksheet name and its "Fit to pages" status
                Console.WriteLine($"Worksheet: {sheet.Name}, FitToPages: {isFitToPages}");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
