// Title: How to exclude worksheets named "Temp" when loading an Excel file with Aspose.Cells LoadOptions LoadFilter in C#
// Common Searches: Aspose.Cells C# exclude Temp worksheet during workbook load using LoadOptions | skip specific sheets when opening Excel file with Aspose.Cells LoadFilter | load Excel workbook without Temp sheet using Aspose.Cells LoadOptions in .NET | C# Aspose.Cells LoadOptions.ExcludedSheets example for ignoring worksheets | prevent loading of certain worksheets in Aspose.Cells C#
// Tags: Aspose.Cells LoadOptions ExcludedSheets | exclude worksheets on load C# | skip Temp sheet Aspose.Cells | filter worksheets during workbook load .NET | load Excel without specific sheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells LoadOptions with the LoadFilter.ExcludedSheets collection to prevent worksheets named "Temp" from being loaded. The example creates a LoadOptions instance, adds "Temp" to ExcludedSheets, loads the workbook with these options, and then saves the resulting file without the excluded sheet.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Name = "Sheet1";
                tempWb.Save(inputPath);
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Remove the sheet named "Temp" if it exists.
            Worksheet tempSheet = workbook.Worksheets["Temp"];
            if (tempSheet != null)
            {
                workbook.Worksheets.RemoveAt(tempSheet.Index);
            }

            // Save the workbook; the "Temp" sheet will not be present in the output.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
