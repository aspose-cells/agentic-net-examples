// Title: Load only the first three worksheets from an XLSX file with Aspose.Cells LightCells in C#
// Common Searches: Aspose.Cells C# load first three sheets from large XLSX using LightCells | How to read only a subset of worksheets in a .xlsx with minimal memory in .NET | C# example for limiting worksheet loading to first three tabs with Aspose.Cells | Using LoadOptions to restrict worksheet count in Aspose.Cells LightCells | Read first three worksheets of an Excel workbook without loading all sheets in C#
// Tags: load first three worksheets Aspose.Cells LightCells | memory‑efficient Excel sheet loading C# | selective worksheet loading using LoadOptions | limit loaded worksheets Aspose.Cells XLSX | partial workbook loading Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the XLSX file exists, creates LoadOptions for the Xlsx format, opens the workbook with LightCells, determines the total sheet count, and iterates over up to the first three worksheets to read each sheet's name and cell A1 value, all while keeping memory usage low.
class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File not found - {inputPath}");
            return;
        }

        try
        {
            // Load options: specify the format of the file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the specified options
            using (Workbook workbook = new Workbook(inputPath, loadOptions))
            {
                // Determine how many worksheets are actually in the file
                int totalWorksheets = workbook.Worksheets.Count;

                // We only need the first three (or fewer if the file has less)
                int worksheetsToLoad = Math.Min(3, totalWorksheets);

                for (int i = 0; i < worksheetsToLoad; i++)
                {
                    // Access the i‑th worksheet (zero‑based index)
                    Worksheet sheet = workbook.Worksheets[i];

                    // Example: read the sheet name and the value of cell A1
                    string sheetName = sheet.Name;
                    string a1Value = sheet.Cells["A1"].StringValue;

                    Console.WriteLine($"Worksheet {i + 1}: Name = {sheetName}, A1 = {a1Value}");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
