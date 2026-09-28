// Title: Compare paper size, orientation, and margins of two worksheets with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that loads an Excel workbook, selects two worksheets by name, and prints any differences in their PageSetup properties such as PaperSize, Orientation, TopMargin, BottomMargin, LeftMargin, and RightMargin. | Create a reusable C# method that accepts two Worksheet objects and returns a collection of mismatched PageSetup settings—including paper size, orientation, and margins—so the caller can handle the results as needed. | Extend the comparison to also evaluate header and footer text, then generate a formatted console report that lists all PageSetup discrepancies between the two worksheets.
// Common Searches: aspnet compare worksheet page layout properties Aspose.Cells | c# detect paper size differences between two Excel sheets using Aspose | how to find margin mismatches in Excel worksheets with Aspose.Cells .NET | retrieve orientation setting from a worksheet with Aspose.Cells C# | log page layout discrepancies between two worksheets in a console application
// Tags: worksheet page layout comparison Aspose.Cells | paper size mismatch detection .NET | margin variance analysis C# | page orientation discrepancy check Aspose | console report of worksheet layout differences

using Aspose.Cells;
using System;

// The program loads an Excel workbook, accesses two worksheets by name, retrieves their PageSetup objects, and compares PaperSize, Orientation, and each margin (Top, Bottom, Left, Right). Any differences are written to the console with a small tolerance for margin values.
class Program
{
    static void Main()
    {
        // Load the workbook containing the worksheets to compare
        var workbook = new Workbook("input.xlsx");

        // Names of the worksheets to compare
        string firstSheetName = "Sheet1";
        string secondSheetName = "Sheet2";

        // Retrieve the worksheets
        var firstSheet = workbook.Worksheets[firstSheetName];
        var secondSheet = workbook.Worksheets[secondSheetName];

        if (firstSheet == null || secondSheet == null)
        {
            Console.WriteLine("One or both specified worksheets were not found.");
            return;
        }

        // Access the PageSetup objects which hold paper dimensions
        var firstPageSetup = firstSheet.PageSetup;
        var secondPageSetup = secondSheet.PageSetup;

        // Compare PaperSize
        if (firstPageSetup.PaperSize != secondPageSetup.PaperSize)
        {
            Console.WriteLine($"PaperSize differs: {firstSheetName} = {firstPageSetup.PaperSize}, {secondSheetName} = {secondPageSetup.PaperSize}");
        }

        // Compare Orientation
        if (firstPageSetup.Orientation != secondPageSetup.Orientation)
        {
            Console.WriteLine($"Orientation differs: {firstSheetName} = {firstPageSetup.Orientation}, {secondSheetName} = {secondPageSetup.Orientation}");
        }

        // Compare margins (values are in points)
        CompareMargin("TopMargin", firstPageSetup.TopMargin, secondPageSetup.TopMargin, firstSheetName, secondSheetName);
        CompareMargin("BottomMargin", firstPageSetup.BottomMargin, secondPageSetup.BottomMargin, firstSheetName, secondSheetName);
        CompareMargin("LeftMargin", firstPageSetup.LeftMargin, secondPageSetup.LeftMargin, firstSheetName, secondSheetName);
        CompareMargin("RightMargin", firstPageSetup.RightMargin, secondPageSetup.RightMargin, firstSheetName, secondSheetName);
    }

    // Helper method to compare individual margin values with a small tolerance
    static void CompareMargin(string marginName, double firstValue, double secondValue, string firstSheet, string secondSheet)
    {
        const double tolerance = 0.001; // points tolerance
        if (Math.Abs(firstValue - secondValue) > tolerance)
        {
            Console.WriteLine($"{marginName} differs: {firstSheet} = {firstValue}, {secondSheet} = {secondValue}");
        }
    }
}
