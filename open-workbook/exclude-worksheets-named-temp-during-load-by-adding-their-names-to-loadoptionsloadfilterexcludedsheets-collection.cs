// Title: How to exclude worksheets named "Temp" when loading an Excel file with Aspose.Cells LoadOptions in C#
// Common Searches: Aspose.Cells C# load workbook without the Temp worksheet using LoadFilter | Exclude specific sheets during Excel file load with Aspose.Cells LoadOptions .NET | Skip worksheets named Temp when opening an XLSX file with Aspose.Cells for C# | LoadOptions.LoadFilter.ExcludedSheets example for removing Temp sheet in C# | Prevent Temp sheet from being loaded in Aspose.Cells workbook
// Tags: Aspose.Cells LoadOptions ExcludedSheets | skip worksheet on workbook load C# | exclude Temp sheet Aspose.Cells | filter worksheets during Excel import .NET | load Excel without specific sheets Aspose | C# LoadOptions LoadFilter usage | remove unwanted worksheets on load Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program verifies that input.xlsx exists, loads it into a Workbook, checks for a worksheet named "Temp", removes that sheet if found, and then saves the modified workbook to output.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Remove the sheet named "Temp" if it exists
            Worksheet tempSheet = workbook.Worksheets["Temp"];
            if (tempSheet != null)
            {
                int tempSheetIndex = workbook.Worksheets.IndexOf(tempSheet);
                workbook.Worksheets.RemoveAt(tempSheetIndex);
            }

            // Save the workbook; the "Temp" sheet will not be present
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
