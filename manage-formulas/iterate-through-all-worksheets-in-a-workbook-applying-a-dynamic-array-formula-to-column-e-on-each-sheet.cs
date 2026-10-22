// Title: Set a SORT dynamic array formula in column E of every worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply the same formula to all worksheets in a workbook | How to assign a dynamic array SORT formula to cell E1 across multiple sheets with Aspose.Cells | Iterate through Excel worksheets and set a formula programmatically in C# | Bulk insert a SORT(A1:D100) formula into column E using Aspose.Cells .NET | Set a formula for each sheet in an existing workbook with Aspose.Cells
// Tags: Aspose.Cells bulk formula assignment | C# iterate workbook worksheets | apply SORT function to Excel column via Aspose | set cell E1 formula across multiple sheets | dynamic array support in Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook (or creates one with sample data), defines the dynamic array formula "=SORT(A1:D100)", loops through all worksheets, writes that formula to cell E1 on each sheet, and saves the workbook, handling load, formula, and save errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        Workbook workbook = null;

        // Load existing workbook if it exists; otherwise create a new one with sample data.
        if (File.Exists(inputPath))
        {
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Error loading workbook '{inputPath}': {ex.Message}");
                return;
            }
        }
        else
        {
            // Create a new workbook and populate A1:D100 with sample values.
            workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 0; row < 100; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    sheet.Cells[row, col].PutValue(row * 4 + col + 1);
                }
            }
        }

        // Define the dynamic array formula to apply.
        string dynamicArrayFormula = "=SORT(A1:D100)";

        // Apply the formula to each worksheet.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            try
            {
                sheet.Cells["E1"].Formula = dynamicArrayFormula;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to set formula on sheet '{sheet.Name}': {ex.Message}");
            }
        }

        // Save the workbook.
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
