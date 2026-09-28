// Title: Cut cells B2:C4 from the first worksheet and paste into G5:H7 on another worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that cuts the range B2:C4 from the first worksheet of a workbook and pastes it into G5:H7 on a second worksheet, creating the destination sheet if it does not exist, using Aspose.Cells. | Show how to move a cell block between worksheets with Aspose.Cells for .NET, including clearing the source cells after copying.
// Common Searches: Aspose.Cells how to cut a range and paste it into another sheet in C# | C# move cells B2:C4 to G5:H7 on a different worksheet using Aspose.Cells | Copy and clear source range with Aspose.Cells .NET example | Create destination worksheet if missing when transferring range in Aspose.Cells
// Tags: cut range between worksheets Aspose.Cells | copy range then clear source Aspose.Cells .NET | transfer cell block B2:C4 to G5:H7 Aspose.Cells | auto-create worksheet Aspose.Cells when moving range | Aspose.Cells range manipulation example

using System;
using System.IO;
using Aspose.Cells;

// The program loads "source.xlsx", cuts the range B2:C4 from the first worksheet, pastes it into G5:H7 on a second worksheet (creating the sheet if needed), clears the original cells, and saves the result as "output.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the source worksheet (first sheet)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Ensure the destination worksheet exists; create it if necessary
            Worksheet destSheet;
            if (workbook.Worksheets.Count > 1)
            {
                destSheet = workbook.Worksheets[1];
            }
            else
            {
                // Add a new worksheet and retrieve it by index
                int newIndex = workbook.Worksheets.Add();
                destSheet = workbook.Worksheets[newIndex];
            }

            // Define the range to cut from the source sheet (B2:C4)
            Aspose.Cells.Range sourceRange = sourceSheet.Cells.CreateRange("B2", "C4");

            // Define the destination range on the destination sheet (G5:H7)
            Aspose.Cells.Range destinationRange = destSheet.Cells.CreateRange("G5", "H7");

            // Copy the source range to the destination range
            sourceRange.Copy(destinationRange);
            // Clear the original cells (simulate a cut operation)
            sourceRange.ClearContents();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
