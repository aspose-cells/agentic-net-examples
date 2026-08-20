// Title: Aspose.Cells .NET: Create a Named Range that Includes Merged Cells
// Description: Demonstrates how to merge cells A1:C2, define a Range covering the merged block, assign a name, set the RefersTo formula correctly, verify the address, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells named range merged cells | C# create named range Aspose | merged cell range RefersTo | Aspose.Range CreateRange example | Workbook.Names add merged area
// Common Searches: how to add a named range that contains merged cells asp.net | aspocells named range reference merged area | c# aspose create range for merged cells | set RefersTo for merged cells Aspose.Cells | verify named range address after merging cells
// Developer Intent: Define a named range covering merged cells and ensure its RefersTo address stays accurate.
// Use Cases: Reference a merged header block in formulas without manual address updates. | Programmatically retrieve and validate a named range that spans merged cells. | Export workbooks where named ranges must point to merged regions for downstream analytics.
// AI Prompts: Write C# code with Aspose.Cells that merges A1:C2, creates a named range for the merged block, and sets RefersTo correctly. | Explain how Aspose.Cells treats merged cells when creating a Range and how to confirm the named range address. | Show how to adjust the RefersTo formula if the merged area is expanded after the named range is created.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNamedRangeWithMergedCells
{
    // Demonstrates how to merge cells A1:C2, define a Range covering the merged block, assign a name, set the RefersTo formula correctly, verify the address, and save the workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Merge cells A1:C2 (rows 0‑1, columns 0‑2)
                sheet.Cells.Merge(0, 0, 2, 3);

                // Put a value into the merged cell (upper‑left cell of the merged area)
                sheet.Cells[0, 0].PutValue("Merged Area");

                // Create a Range object that covers the merged cells
                AsposeRange mergedRange = sheet.Cells.CreateRange("A1", "C2");

                // Assign a name to the range (named range)
                mergedRange.Name = "MyMergedRange";

                // Add the named range to the workbook's Names collection
                int nameIndex = workbook.Worksheets.Names.Add("MyMergedRange");
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                // RefersTo must start with '=' and include the sheet name
                namedRange.RefersTo = $"={sheet.Name}!{mergedRange.RefersTo}";

                // Verify that the reference is correct by retrieving the range via the Name object
                AsposeRange retrievedRange = namedRange.GetRange();
                Console.WriteLine("Named range address: " + retrievedRange.Address);
                Console.WriteLine("Is the address the same as the original? " +
                                  (retrievedRange.Address == mergedRange.Address));

                // Save the workbook
                string outputPath = "NamedRangeWithMergedCells.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
