// Title: Configure the right page margin to 10 points in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that sets the worksheet's right margin to 10 points via Aspose.Cells PageSetup. | Create an Excel workbook where the right margin is fixed at 10 points to accommodate long sentences, using Aspose.Cells in .NET. | Write a snippet that modifies the PageSetup.RightMargin property to a specific point value with Aspose.Cells.
// Common Searches: Aspose.Cells C# set worksheet right margin in points | how to change right page margin to 10 points using Aspose.Cells .NET | adjust Excel worksheet margins by points with Aspose.Cells API | set right margin for long text in Excel using Aspose.Cells C# | PageSetup RightMargin property Aspose.Cells example
// Tags: Aspose.Cells right margin points | Excel page setup margin adjustment C# | configure worksheet right margin Aspose.Cells | long sentence margin handling .NET | PageSetup RightMargin property example

using Aspose.Cells;

// The example creates a new workbook, accesses the first worksheet, sets the PageSetup.RightMargin to 10 points, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the right margin to ten points
        sheet.PageSetup.RightMargin = 10f; // points

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
