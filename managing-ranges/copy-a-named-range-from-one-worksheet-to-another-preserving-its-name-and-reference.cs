// Title: Copy a Named Range Between Workbooks with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to copy a defined named range (including its data, formatting, and formula reference) from a source worksheet to a destination worksheet in a different workbook, while preserving the original name and updating the RefersTo address.
// Keywords: Aspose.Cells copy named range | C# copy range between workbooks | preserve named range RefersTo | duplicate named range Aspose.Cells | copy range formatting .NET
// Common Searches: copy named range Aspose.Cells C# | preserve RefersTo when moving range to another workbook | duplicate named range across workbooks .NET | Aspose.Cells copy range with formatting | how to replicate a named range in a new workbook
// Developer Intent: Transfer a named range from one workbook to another, recreate the same name in the target workbook, and adjust the RefersTo formula to point to the new location.
// Use Cases: Migrate a template range with formulas from a master file to client‑specific files while keeping the range identifier. | Propagate data‑validation or calculation ranges across multiple reports generated programmatically. | Synchronize named range definitions when consolidating data from several source workbooks.
// AI Prompts: Generate C# code using Aspose.Cells to copy a named range from one workbook to another and update its RefersTo reference. | Explain step‑by‑step how to duplicate a named range, including data, formatting, and name, between worksheets in .NET. | Provide robust error‑handling patterns for copying named ranges across workbooks with Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to copy a defined named range (including its data, formatting, and formula reference) from a source worksheet to a destination worksheet in a different workbook, while preserving the original name and updating the RefersTo address.
class CopyNamedRangeDemo
{
    static void Main()
    {
        try
        {
            // ---------- Create source workbook ----------
            Workbook srcWb = new Workbook();                     // create source workbook
            Worksheet srcWs = srcWb.Worksheets[0];
            srcWs.Name = "Source";

            // Fill some data in the source range
            srcWs.Cells["A1"].PutValue("Item");
            srcWs.Cells["B1"].PutValue(10);
            srcWs.Cells["A2"].PutValue("Qty");
            srcWs.Cells["B2"].PutValue(20);

            // Define a named range "MyRange" that refers to A1:B2 on the source sheet
            int srcNameIdx = srcWb.Worksheets.Names.Add("MyRange");
            srcWb.Worksheets.Names[srcNameIdx].RefersTo = $"={srcWs.Name}!$A$1:$B$2";

            // ---------- Create destination workbook ----------
            Workbook destWb = new Workbook();                    // create destination workbook
            Worksheet destWs = destWb.Worksheets[0];
            destWs.Name = "Destination";

            // ---------- Copy the cells of the named range ----------
            // Retrieve the source range via the Name object
            Name srcName = srcWb.Worksheets.Names["MyRange"];
            AsposeRange srcRange = srcName.GetRange();

            // Create an equally sized range on the destination sheet at the same address
            AsposeRange destRange = destWs.Cells.CreateRange(
                srcRange.FirstRow,
                srcRange.FirstColumn,
                srcRange.RowCount,
                srcRange.ColumnCount);

            // Copy data, formatting, etc. from source range to destination range
            destRange.Copy(srcRange);

            // ---------- Replicate the named range in the destination workbook ----------
            int destNameIdx = destWb.Worksheets.Names.Add("MyRange");
            // Adjust the RefersTo formula to point to the destination sheet
            destWb.Worksheets.Names[destNameIdx].RefersTo = $"={destWs.Name}!$A$1:$B$2";

            // ---------- Save workbooks ----------
            srcWb.Save("Source.xlsx");
            destWb.Save("Destination.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
