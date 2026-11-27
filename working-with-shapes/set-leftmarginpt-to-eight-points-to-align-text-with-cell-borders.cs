// Title: How to set the worksheet left page margin to 8 points with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet left margin to 8 points | page setup left margin points example using Aspose.Cells .NET | adjust Excel left page margin programmatically with Aspose.Cells | align cell text with borders by changing left margin in Aspose.Cells | set left margin in points for Excel workbook using Aspose.Cells C#
// Tags: set left margin points Aspose.Cells | worksheet page setup left margin C# | adjust Excel page margins programmatically | align text with cell borders Aspose.Cells | Aspose.Cells page setup margin configuration

using Aspose.Cells;
using System;

// The example creates a new workbook, accesses the first worksheet, sets the page setup left margin to 8 points to align text with cell borders, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample text to a cell (optional, just for context)
        sheet.Cells["A1"].PutValue("Sample text");

        // Set the left margin to eight points to align text with cell borders
        sheet.PageSetup.LeftMargin = 8; // margin in points

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
