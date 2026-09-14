// Title: Check each worksheet for dialog sheet type in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells and prints the name and dialog‑sheet status of every worksheet. | Show how to use the Worksheet.Type property together with SheetType.Dialog to filter dialog sheets in a .NET application. | Write a reusable C# method that returns a list of worksheet names that are dialog sheets from a given workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to determine if a worksheet is a dialog sheet | C# code sample for detecting dialog sheets in an Excel workbook with Aspose.Cells | Using SheetType.Dialog enumeration to list dialog worksheets in .NET | Identify dialog sheet type in Excel file using Aspose.Cells library
// Tags: Aspose.Cells worksheet type detection | C# SheetType.Dialog usage | list dialog sheets Aspose.Cells | Excel dialog sheet identification .NET | filter worksheets by type Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads 'input.xlsx' with Aspose.Cells, iterates through all worksheets, checks each worksheet's Type against SheetType.Dialog, and prints the worksheet name along with a boolean indicating whether it is a dialog sheet, while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine whether the worksheet is a dialog sheet
                bool isDialog = sheet.Type == SheetType.Dialog;

                // Output the result
                Console.WriteLine($"Worksheet \"{sheet.Name}\" is a dialog sheet: {isDialog}");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
