// Title: Copy a Named Range with Absolute References Between Worksheets using Aspose.Cells for .NET
// Description: Demonstrates how to create a source workbook, define a named range with $A$1 style absolute references, and duplicate that range on a different worksheet while preserving values, formulas, and formatting. The example also shows how to recreate the Name object in the target workbook and save the result as an Excel file.
// Keywords: Aspose.Cells copy named range | absolute cell references C# | duplicate named range workbook | Aspose.Cells .NET range copy | preserve formulas Aspose.Cells | Excel named range API
// Common Searches: Aspose.Cells copy named range to another sheet | preserve $A$1 references when copying range .NET | duplicate named range across workbooks C# | how to copy range with formulas using Aspose.Cells | copy named range between worksheets programmatically
// Developer Intent: Duplicate a named range that uses absolute cell references from one worksheet to another without altering the references.
// Use Cases: Reuse a predefined data block in multiple report tabs while keeping formula links intact. | Generate department‑specific templates that share the same named range and calculations. | Export a specific named range from a master workbook to a client‑facing file without breaking references.
// AI Prompts: Write C# code with Aspose.Cells that copies a named range containing absolute references from a source worksheet to a destination worksheet and updates the RefersTo property. | Show how to retrieve a Name object's Range, copy its contents to another sheet, and recreate the same Name in the target workbook using Aspose.Cells for .NET. | Explain the steps to preserve formulas, formatting, and $A$1‑style references when moving a named range between Excel worksheets programmatically.

using System;
using Aspose.Cells;

// Demonstrates how to create a source workbook, define a named range with $A$1 style absolute references, and duplicate that range on a different worksheet while preserving values, formulas, and formatting. The example also shows how to recreate the Name object in the target workbook and save the result as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create source workbook and define a named range ----------
            Workbook srcWb = new Workbook();
            Worksheet srcSheet = srcWb.Worksheets[0];
            srcSheet.Name = "Source";

            // Fill some data in the source range
            srcSheet.Cells["A1"].PutValue(10);
            srcSheet.Cells["A2"].PutValue(20);
            srcSheet.Cells["B1"].PutValue(30);
            srcSheet.Cells["B2"].PutValue(40);

            // Add a named range with absolute references (e.g., $A$1:$B$2)
            int srcNameIdx = srcWb.Worksheets.Names.Add("MyRange");
            Name srcName = srcWb.Worksheets.Names[srcNameIdx];
            srcName.RefersTo = $"={srcSheet.Name}!$A$1:$B$2";

            // ---------- Create destination workbook ----------
            Workbook destWb = new Workbook();
            Worksheet destSheet = destWb.Worksheets[0];
            destSheet.Name = "Destination";

            // ---------- Retrieve the source range via the Name object ----------
            // GetRange() returns the actual Range object the name refers to
            Aspose.Cells.Range srcRange = srcName.GetRange();

            // ---------- Create a destination range of the same size ----------
            // Here we copy to the same address in the destination sheet
            Aspose.Cells.Range destRange = destSheet.Cells.CreateRange("A1:B2");

            // ---------- Copy the source range to the destination range ----------
            // This copies values, formulas, formatting, etc., preserving absolute references
            destRange.Copy(srcRange);

            // ---------- Replicate the named range in the destination workbook ----------
            // The RefersTo string must point to the destination sheet
            int destNameIdx = destWb.Worksheets.Names.Add("MyRange");
            Name destName = destWb.Worksheets.Names[destNameIdx];
            destName.RefersTo = $"={destSheet.Name}!$A$1:$B$2";

            // ---------- Save the result ----------
            destWb.Save("CopyNamedRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
