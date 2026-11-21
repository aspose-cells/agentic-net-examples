// Title: Set an Excel slicer to be printable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set slicer printable before saving workbook | Make Excel slicer appear on printed page using Aspose.Cells .NET | How to enable slicer printing in Aspose.Cells for .NET | C# code to set IsPrintable property of a slicer in Aspose.Cells | Include slicer in printed worksheet with Aspose.Cells library
// Tags: Aspose.Cells slicer IsPrintable | C# set slicer printable Aspose.Cells | Excel slicer print visibility .NET | Aspose.Cells modify slicer print flag | save workbook with printable slicer Aspose

using Aspose.Cells;

// The sample loads an existing workbook, checks for slicers on the first worksheet, sets the first slicer's IsPrintable property to true, and saves the workbook so the slicer is included when the worksheet is printed.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Check that the worksheet contains at least one slicer
        if (sheet.Slicers.Count > 0)
        {
            // Set the first slicer to be printable so it appears when printed
            sheet.Slicers[0].IsPrintable = true;
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
