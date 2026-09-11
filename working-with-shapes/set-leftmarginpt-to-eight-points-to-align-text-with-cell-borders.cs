// Title: Set worksheet left page margin to 8 points using Aspose.Cells for .NET (C#)
// AI Prompts: Apply Aspose.Cells PageSetup.LeftMargin to set an 8‑point left margin for a worksheet in C#. | Programmatically align Excel cell text with borders by configuring the left margin to 8 points via Aspose.Cells. | Create a workbook, adjust the left page margin to 8 points, and save the file using Aspose.Cells in .NET.
// Common Searches: Aspose.Cells C# set worksheet left margin 8 points | how to change page left margin in Excel using Aspose.Cells .NET | align text with cell borders by adjusting left margin in Aspose.Cells workbook | PageSetup.LeftMargin property example Aspose.Cells C# | set Excel page margins programmatically Aspose.Cells
// Tags: Aspose.Cells PageSetup.LeftMargin property | C# set worksheet left margin points | Excel workbook left page margin adjustment | align cell text with borders Aspose.Cells | programmatic Excel page margin configuration .NET

using Aspose.Cells;

// The example creates a new Workbook, accesses the first Worksheet, sets the left page margin to 8 points via the PageSetup.LeftMargin property, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the left margin to 8 points to align text with cell borders
        sheet.PageSetup.LeftMargin = 8f;

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
