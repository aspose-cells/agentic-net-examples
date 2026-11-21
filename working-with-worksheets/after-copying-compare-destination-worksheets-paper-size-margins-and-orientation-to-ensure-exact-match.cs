// Title: Verify that a copied worksheet retains the same paper size, margins, and orientation using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells compare page setup of source and copied worksheet | how to check if worksheet margins are identical after copying with Aspose.Cells | validate paper size and orientation after worksheet copy in .NET | detect mismatched page setup properties between two worksheets using Aspose.Cells | Aspose.Cells C# ensure copied sheet keeps original page layout settings
// Tags: Aspose.Cells worksheet page setup comparison | C# verify worksheet margins after copy | Aspose.Cells paper size orientation check | page setup property validation Aspose.Cells | copy worksheet preserve page layout .NET

using Aspose.Cells;
using System;

// The example loads a source Excel workbook, copies its first worksheet into a new workbook, and then retrieves the PageSetup objects for both the source and destination worksheets. It compares paper size, orientation, and all margin values, logs any mismatches, reports whether the page setup matches exactly, and finally saves the destination workbook.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook sourceWorkbook = new Workbook("source.xlsx");

        // Create an empty destination workbook
        Workbook destinationWorkbook = new Workbook();

        // Copy the first worksheet from source to destination
        // (Assumes at least one worksheet exists in the source)
        sourceWorkbook.Worksheets[0].Copy(destinationWorkbook.Worksheets[0]);

        // Get references to the source and destination worksheets
        Worksheet sourceWorksheet = sourceWorkbook.Worksheets[0];
        Worksheet destinationWorksheet = destinationWorkbook.Worksheets[0];

        // Retrieve the PageSetup objects for both worksheets
        PageSetup sourceSetup = sourceWorksheet.PageSetup;
        PageSetup destinationSetup = destinationWorksheet.PageSetup;

        // Flag to track if all settings match
        bool isExactMatch = true;

        // Compare paper size
        if (sourceSetup.PaperSize != destinationSetup.PaperSize)
        {
            isExactMatch = false;
            Console.WriteLine($"PaperSize mismatch: Source={sourceSetup.PaperSize}, Destination={destinationSetup.PaperSize}");
        }

        // Compare orientation
        if (sourceSetup.Orientation != destinationSetup.Orientation)
        {
            isExactMatch = false;
            Console.WriteLine($"Orientation mismatch: Source={sourceSetup.Orientation}, Destination={destinationSetup.Orientation}");
        }

        // Compare margins (values are in points)
        if (sourceSetup.TopMargin != destinationSetup.TopMargin)
        {
            isExactMatch = false;
            Console.WriteLine($"TopMargin mismatch: Source={sourceSetup.TopMargin}, Destination={destinationSetup.TopMargin}");
        }

        if (sourceSetup.BottomMargin != destinationSetup.BottomMargin)
        {
            isExactMatch = false;
            Console.WriteLine($"BottomMargin mismatch: Source={sourceSetup.BottomMargin}, Destination={destinationSetup.BottomMargin}");
        }

        if (sourceSetup.LeftMargin != destinationSetup.LeftMargin)
        {
            isExactMatch = false;
            Console.WriteLine($"LeftMargin mismatch: Source={sourceSetup.LeftMargin}, Destination={destinationSetup.LeftMargin}");
        }

        if (sourceSetup.RightMargin != destinationSetup.RightMargin)
        {
            isExactMatch = false;
            Console.WriteLine($"RightMargin mismatch: Source={sourceSetup.RightMargin}, Destination={destinationSetup.RightMargin}");
        }

        if (sourceSetup.HeaderMargin != destinationSetup.HeaderMargin)
        {
            isExactMatch = false;
            Console.WriteLine($"HeaderMargin mismatch: Source={sourceSetup.HeaderMargin}, Destination={destinationSetup.HeaderMargin}");
        }

        if (sourceSetup.FooterMargin != destinationSetup.FooterMargin)
        {
            isExactMatch = false;
            Console.WriteLine($"FooterMargin mismatch: Source={sourceSetup.FooterMargin}, Destination={destinationSetup.FooterMargin}");
        }

        // Output final result
        if (isExactMatch)
        {
            Console.WriteLine("Destination worksheet's paper size, margins, and orientation exactly match the source.");
        }
        else
        {
            Console.WriteLine("There are mismatches in the destination worksheet's page setup.");
        }

        // Save the destination workbook
        destinationWorkbook.Save("destination.xlsx");
    }
}
