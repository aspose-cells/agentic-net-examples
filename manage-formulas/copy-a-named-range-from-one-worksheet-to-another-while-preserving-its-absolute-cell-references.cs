// Title: Copy a Named Range with Absolute References Between Worksheets – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to copy a named range that uses absolute cell references (e.g., $A$1:$B$2) from one worksheet to another in a workbook created with Aspose.Cells for .NET. The example shows creating the source range, retrieving it via the name, defining a destination range with CreateRange, using Range.Copy to preserve the $ signs, assigning a new name, and saving the file.
// Keywords: Aspose.Cells copy named range | absolute cell references | C# Excel automation | Range.Copy Aspose.Cells | GetRange Aspose.Cells | CreateRange worksheet | RefersTo property | duplicate named range .NET | preserve $ signs Excel | copy range between sheets
// Common Searches: copy named range to another sheet Aspose.Cells C# | preserve absolute references when copying range | Aspose.Cells Range.Copy example | how to duplicate a named range in .NET | retain $ signs in copied Excel range Aspose
// Developer Intent: Copy an existing named range to a different worksheet while keeping its absolute cell references unchanged.
// Use Cases: Create a summary tab that mirrors a fixed data block from a source sheet without breaking formulas. | Generate multiple template sections that require the same absolute range for consistent reporting. | Duplicate a named range for a new scenario, assign a distinct name, and export the workbook for downstream processing.
// AI Prompts: Show C# code that copies a named range with $-style absolute references from one worksheet to another using Aspose.Cells. | Explain how to retrieve a named range, copy it to a new location, and assign a new name while preserving the RefersTo $ signs. | Provide a step‑by‑step Aspose.Cells example for duplicating a named range across worksheets and updating the RefersTo property.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to copy a named range that uses absolute cell references (e.g., $A$1:$B$2) from one worksheet to another in a workbook created with Aspose.Cells for .NET. The example shows creating the source range, retrieving it via the name, defining a destination range with CreateRange, using Range.Copy to preserve the $ signs, assigning a new name, and saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------- Source Worksheet --------------------
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Populate some data in the source sheet
            sourceSheet.Cells["A1"].PutValue(10);
            sourceSheet.Cells["A2"].PutValue(20);
            sourceSheet.Cells["B1"].PutValue(30);
            sourceSheet.Cells["B2"].PutValue(40);

            // Define a named range that uses absolute references ($A$1:$B$2)
            int srcNameIdx = workbook.Worksheets.Names.Add("MyRange");
            workbook.Worksheets.Names[srcNameIdx].RefersTo = $"={sourceSheet.Name}!$A$1:$B$2";

            // -------------------- Destination Worksheet --------------------
            Worksheet destSheet = workbook.Worksheets.Add("Destination");

            // Retrieve the source range via the name (uses Name.GetRange rule)
            AsposeRange sourceRange = workbook.Worksheets.Names["MyRange"].GetRange();

            // Create a destination range of the same size (C3:D4 in this example)
            AsposeRange destRange = destSheet.Cells.CreateRange("C3:D4");

            // Copy the source range to the destination range while preserving absolute references
            sourceRange.Copy(destRange);

            // Create a new name for the copied range (optional, shows the range is now available under a new name)
            int destNameIdx = workbook.Worksheets.Names.Add("MyRangeCopy");
            workbook.Worksheets.Names[destNameIdx].RefersTo = $"={destSheet.Name}!$C$3:$D$4";

            // Save the workbook
            workbook.Save("NamedRangeCopy.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
