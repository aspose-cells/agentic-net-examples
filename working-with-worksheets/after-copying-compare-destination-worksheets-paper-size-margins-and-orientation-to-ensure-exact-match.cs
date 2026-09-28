// Title: Verify that a worksheet copied with Aspose.Cells retains identical paper size, orientation, and margins in C#
// AI Prompts: Write C# code using Aspose.Cells to copy a worksheet and then programmatically compare the source and destination PageSetup properties (PaperSize, Orientation, LeftMargin, RightMargin, TopMargin, BottomMargin) with a tolerance for margin values. | Create a reusable method that accepts two Worksheet objects and returns a detailed report indicating which page‑setup attributes match or differ, suitable for unit‑testing worksheet duplication.
// Common Searches: C# Aspose.Cells how to compare page setup of two worksheets after copy | check if copied Excel sheet keeps original paper size and margins using Aspose.Cells | Aspose.Cells .NET verify worksheet orientation equality after duplication
// Tags: Aspose.Cells worksheet page setup verification | C# compare worksheet paper size orientation | Aspose.Cells margin tolerance comparison | Excel page setup equality .NET

using System;
using Aspose.Cells;

// The example loads a source workbook, copies its first worksheet into a new workbook, accesses the PageSetup objects of both worksheets, and compares PaperSize, Orientation, and each margin (left, right, top, bottom) using a small tolerance for floating‑point differences. It prints the comparison results and saves the destination workbook.
class WorksheetPageSetupComparer
{
    static void Main()
    {
        // Load the source workbook
        Workbook srcWorkbook = new Workbook("Source.xlsx");

        // Assume we are working with the first worksheet as source
        Worksheet srcSheet = srcWorkbook.Worksheets[0];

        // Create a new workbook for the destination (or load an existing one)
        Workbook destWorkbook = new Workbook();
        // Add a new worksheet to the destination workbook
        Worksheet destSheet = destWorkbook.Worksheets[destWorkbook.Worksheets.Add()];
        // Copy the source worksheet content to the destination worksheet
        destSheet.Copy(srcSheet);

        // Access PageSetup objects for both worksheets
        PageSetup srcSetup = srcSheet.PageSetup;
        PageSetup destSetup = destSheet.PageSetup;

        // Compare PaperSize
        bool paperSizeMatch = srcSetup.PaperSize == destSetup.PaperSize;

        // Compare Orientation
        bool orientationMatch = srcSetup.Orientation == destSetup.Orientation;

        // Compare Margins (in points)
        const double tolerance = 0.001; // tolerance for floating point comparison
        bool leftMarginMatch = Math.Abs(srcSetup.LeftMargin - destSetup.LeftMargin) < tolerance;
        bool rightMarginMatch = Math.Abs(srcSetup.RightMargin - destSetup.RightMargin) < tolerance;
        bool topMarginMatch = Math.Abs(srcSetup.TopMargin - destSetup.TopMargin) < tolerance;
        bool bottomMarginMatch = Math.Abs(srcSetup.BottomMargin - destSetup.BottomMargin) < tolerance;

        // Output comparison results
        Console.WriteLine($"Paper Size Match: {paperSizeMatch}");
        Console.WriteLine($"Orientation Match: {orientationMatch}");
        Console.WriteLine($"Left Margin Match: {leftMarginMatch}");
        Console.WriteLine($"Right Margin Match: {rightMarginMatch}");
        Console.WriteLine($"Top Margin Match: {topMarginMatch}");
        Console.WriteLine($"Bottom Margin Match: {bottomMarginMatch}");

        // Save the destination workbook if needed
        destWorkbook.Save("Destination.xlsx");
    }
}
