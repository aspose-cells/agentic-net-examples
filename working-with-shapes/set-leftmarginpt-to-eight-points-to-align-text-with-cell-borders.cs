// Title: Set the worksheet left page margin to 8 points with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells PageSetup to set the left margin of a worksheet to 8 points in C#. | Adjust the worksheet's left page margin so text aligns with cell borders using Aspose.Cells. | Create a workbook, configure its left margin to 8 points, and save it as an .xlsx file with Aspose.Cells.
// Common Searches: Aspose.Cells C# set left margin to 8 points | How to align Excel text with cell borders by adjusting page margins in Aspose.Cells | PageSetup.LeftMargin property example in Aspose.Cells for .NET | Set worksheet page margins programmatically using Aspose.Cells C# | MarginExample.xlsx left margin 8 points Aspose.Cells
// Tags: worksheet left margin configuration Aspose.Cells | PageSetup.LeftMargin property usage | programmatic Excel page margin adjustment | text alignment with cell borders via margins | saving workbook with custom page margins

using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, sets its PageSetup.LeftMargin to 8 points to align text with cell borders, and saves the file as MarginExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        var workbook = new Workbook();

        // Get the first worksheet
        var worksheet = workbook.Worksheets[0];

        // Set the left margin to eight points to align text with cell borders
        worksheet.PageSetup.LeftMargin = 8.0; // points

        // Save the workbook (lifecycle rule)
        workbook.Save("MarginExample.xlsx");
    }
}
