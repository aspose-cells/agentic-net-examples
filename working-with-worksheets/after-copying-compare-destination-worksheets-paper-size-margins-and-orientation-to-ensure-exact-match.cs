// Title: Validate that a copied worksheet preserves paper size, orientation, and margins using Aspose.Cells for .NET
// Description: C# example that creates a source workbook, sets A3 paper size, landscape orientation, and custom margins, copies the sheet with Worksheet.Copy, then programmatically compares the destination's page‑setup properties to the source using enum equality and a tolerance for floating‑point margins. Outputs the verification result and saves both workbooks.
// Keywords: Aspose.Cells copy worksheet page setup | C# verify paper size after copy | worksheet orientation comparison Aspose.Cells | margin tolerance Aspose.Cells .NET | page setup validation after Worksheet.Copy | Aspose.Cells print settings preservation | compare source and destination worksheet layout
// Common Searches: Aspose.Cells check page setup after worksheet copy | C# compare margins of two worksheets | verify paper size orientation after copying sheet Aspose.Cells | how to ensure print settings are retained in copied worksheet | Aspose.Cells Worksheet.Copy page layout validation
// Developer Intent: Confirm that the destination worksheet’s page‑setup (paper size, orientation, margins) exactly matches the source worksheet after a copy operation.
// Use Cases: Automated testing of template worksheets to guarantee identical print layouts for generated reports. | Batch processing where worksheets are duplicated and must retain original page‑setup to avoid printing errors. | Logging discrepancies in page‑setup after copying to trigger corrective actions in a CI pipeline.
// AI Prompts: Write C# code with Aspose.Cells that copies a worksheet and verifies that paper size, orientation, and all margins are identical, using a small tolerance for margin values. | Refactor the example to extract the page‑setup comparison into a reusable method that returns a detailed mismatch report. | Extend the validation to include header/footer content, scaling, and print area settings after copying a worksheet with Aspose.Cells.

using System;
using Aspose.Cells;

// C# example that creates a source workbook, sets A3 paper size, landscape orientation, and custom margins, copies the sheet with Worksheet.Copy, then programmatically compares the destination's page‑setup properties to the source using enum equality and a tolerance for floating‑point margins. Outputs the verification result and saves both workbooks.
class Program
{
    static void Main()
    {
        // ---------- Create source workbook and set its page setup ----------
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
        sourceSheet.Name = "Source";

        // Configure paper size, orientation and margins (centimeters)
        sourceSheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
        sourceSheet.PageSetup.Orientation = PageOrientationType.Landscape;
        sourceSheet.PageSetup.TopMargin = 2.0;
        sourceSheet.PageSetup.BottomMargin = 2.5;
        sourceSheet.PageSetup.LeftMargin = 1.5;
        sourceSheet.PageSetup.RightMargin = 1.5;

        // ---------- Create destination workbook ----------
        Workbook destWorkbook = new Workbook();
        Worksheet destSheet = destWorkbook.Worksheets[0];
        destSheet.Name = "Destination";

        // ---------- Copy source worksheet into destination worksheet ----------
        // Using Worksheet.Copy with default CopyOptions
        destSheet.Copy(sourceSheet, new CopyOptions());

        // ---------- Compare page setup properties ----------
        bool isExactMatch = true;

        // Paper size
        isExactMatch &= destSheet.PageSetup.PaperSize == sourceSheet.PageSetup.PaperSize;

        // Orientation
        isExactMatch &= destSheet.PageSetup.Orientation == sourceSheet.PageSetup.Orientation;

        // Margins (use a tolerance for floating‑point comparison)
        const double tolerance = 1e-6;
        isExactMatch &= Math.Abs(destSheet.PageSetup.TopMargin - sourceSheet.PageSetup.TopMargin) < tolerance;
        isExactMatch &= Math.Abs(destSheet.PageSetup.BottomMargin - sourceSheet.PageSetup.BottomMargin) < tolerance;
        isExactMatch &= Math.Abs(destSheet.PageSetup.LeftMargin - sourceSheet.PageSetup.LeftMargin) < tolerance;
        isExactMatch &= Math.Abs(destSheet.PageSetup.RightMargin - sourceSheet.PageSetup.RightMargin) < tolerance;

        // Output the comparison result
        Console.WriteLine("Page setup exact match after copy: " + isExactMatch);
        if (!isExactMatch)
        {
            Console.WriteLine("Mismatch details:");
            Console.WriteLine($"PaperSize - Source: {sourceSheet.PageSetup.PaperSize}, Dest: {destSheet.PageSetup.PaperSize}");
            Console.WriteLine($"Orientation - Source: {sourceSheet.PageSetup.Orientation}, Dest: {destSheet.PageSetup.Orientation}");
            Console.WriteLine($"TopMargin - Source: {sourceSheet.PageSetup.TopMargin}, Dest: {destSheet.PageSetup.TopMargin}");
            Console.WriteLine($"BottomMargin - Source: {sourceSheet.PageSetup.BottomMargin}, Dest: {destSheet.PageSetup.BottomMargin}");
            Console.WriteLine($"LeftMargin - Source: {sourceSheet.PageSetup.LeftMargin}, Dest: {destSheet.PageSetup.LeftMargin}");
            Console.WriteLine($"RightMargin - Source: {sourceSheet.PageSetup.RightMargin}, Dest: {destSheet.PageSetup.RightMargin}");
        }

        // ---------- Save workbooks (optional) ----------
        sourceWorkbook.Save("SourceWorkbook.xlsx");
        destWorkbook.Save("DestinationWorkbook.xlsx");
    }
}
