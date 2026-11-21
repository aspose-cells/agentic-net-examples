// Title: Replace an existing named range with a new address and recalculate formulas in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# update named range reference to a new cell address | replace existing named range and recalculate formulas in .xlsx using Aspose.Cells | add missing named range programmatically with Aspose.Cells for .NET | recalculate all formulas after changing a named range in Aspose.Cells workbook | load workbook, modify named range, and save with Aspose.Cells C# example
// Tags: named range address update Aspose.Cells | add new named range Aspose.Cells C# | Workbook.CalculateFormula after range change | modify Name.RefersTo property Aspose.Cells | Excel workbook formula recalculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook (creating a new one if the file does not exist), checks for a named range called "MyRange", updates its RefersTo address to "Sheet1!$C$1:$D$5" or adds the range if it is missing, triggers a full formula recalculation, and saves the result to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Load existing workbook or create a new one if the file is missing.
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    // Ensure there is at least one worksheet named "Sheet1".
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Name = "Sheet1";
                }

                // Name of the range to update or create.
                string oldRangeName = "MyRange";

                // New address for the named range (e.g., Sheet1!$C$1:$D$5).
                string newRangeAddress = "Sheet1!$C$1:$D$5";

                // Retrieve the existing named range, if any.
                Name existingName = workbook.Worksheets.Names[oldRangeName];

                if (existingName != null)
                {
                    // Update the reference of the existing named range.
                    existingName.RefersTo = newRangeAddress;
                }
                else
                {
                    // Add a new named range with the desired address.
                    int index = workbook.Worksheets.Names.Add(oldRangeName);
                    Name newName = workbook.Worksheets.Names[index];
                    newName.RefersTo = newRangeAddress;
                }

                // Recalculate all formulas to reflect the changed range.
                workbook.CalculateFormula();

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
