// Title: Validate that a copied worksheet retains the same page setup (paper size, margins, orientation) using Aspose.Cells for .NET
// AI Prompts: Generate C# code that copies a worksheet with Aspose.Cells and then programmatically compares the source and destination PageSetup properties, logging any differences in paper size, orientation, and all margins. | Create a reusable C# method that takes two Worksheet objects and returns true only if their page setup (paper size, orientation, top/bottom/left/right/header/footer margins) matches, using the Aspose.Cells API. | Write a C# console application that loads an Excel file, copies the first worksheet, verifies that the copy’s print settings are identical to the original, and saves the workbook, employing Aspose.Cells PageSetup features.
// Common Searches: how to compare page setup of two worksheets after copying with Aspose.Cells C# | Aspose.Cells verify that copied sheet has same paper size and margins | C# check orientation consistency between source and destination worksheets in Aspose.Cells | detect mismatched print settings after worksheet duplication using Aspose.Cells | Aspose.Cells page setup property comparison example
// Tags: aspocells worksheet copy page-setup validation | c# compare print settings aspocells | aspocells page-setup property comparison | c# verify paper size and margins aspocells | aspocells detect mismatched orientation after copy

using Aspose.Cells;
using System;

// The example loads 'source.xlsx', copies its first worksheet within the same workbook, and retrieves the PageSetup objects for both the original and the copied sheets. It then compares paper size, orientation, and all margin values (top, bottom, left, right, header, footer), logs any mismatches, reports a full match when all properties are identical, and finally saves the workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook srcWb = new Workbook("source.xlsx");

        // Get the source worksheet (first worksheet)
        Worksheet srcWs = srcWb.Worksheets[0];

        // Copy the source worksheet to a new worksheet within the same workbook
        int copyIndex = srcWb.Worksheets.AddCopy(srcWs.Index);
        Worksheet destWs = srcWb.Worksheets[copyIndex];

        // Retrieve page setup objects for comparison
        PageSetup srcSetup = srcWs.PageSetup;
        PageSetup destSetup = destWs.PageSetup;

        bool isMatch = true;

        // Compare paper size
        if (srcSetup.PaperSize != destSetup.PaperSize)
        {
            isMatch = false;
            Console.WriteLine($"PaperSize mismatch: Source={srcSetup.PaperSize}, Destination={destSetup.PaperSize}");
        }

        // Compare orientation
        if (srcSetup.Orientation != destSetup.Orientation)
        {
            isMatch = false;
            Console.WriteLine($"Orientation mismatch: Source={srcSetup.Orientation}, Destination={destSetup.Orientation}");
        }

        // Compare margins (values are in points)
        if (srcSetup.TopMargin != destSetup.TopMargin)
        {
            isMatch = false;
            Console.WriteLine($"TopMargin mismatch: Source={srcSetup.TopMargin}, Destination={destSetup.TopMargin}");
        }
        if (srcSetup.BottomMargin != destSetup.BottomMargin)
        {
            isMatch = false;
            Console.WriteLine($"BottomMargin mismatch: Source={srcSetup.BottomMargin}, Destination={destSetup.BottomMargin}");
        }
        if (srcSetup.LeftMargin != destSetup.LeftMargin)
        {
            isMatch = false;
            Console.WriteLine($"LeftMargin mismatch: Source={srcSetup.LeftMargin}, Destination={destSetup.LeftMargin}");
        }
        if (srcSetup.RightMargin != destSetup.RightMargin)
        {
            isMatch = false;
            Console.WriteLine($"RightMargin mismatch: Source={srcSetup.RightMargin}, Destination={destSetup.RightMargin}");
        }
        if (srcSetup.HeaderMargin != destSetup.HeaderMargin)
        {
            isMatch = false;
            Console.WriteLine($"HeaderMargin mismatch: Source={srcSetup.HeaderMargin}, Destination={destSetup.HeaderMargin}");
        }
        if (srcSetup.FooterMargin != destSetup.FooterMargin)
        {
            isMatch = false;
            Console.WriteLine($"FooterMargin mismatch: Source={srcSetup.FooterMargin}, Destination={destSetup.FooterMargin}");
        }

        // Output final result
        if (isMatch)
        {
            Console.WriteLine("Destination worksheet's page setup matches the source worksheet exactly.");
        }

        // Save the workbook with the copied worksheet (optional)
        srcWb.Save("output.xlsx");
    }
}
