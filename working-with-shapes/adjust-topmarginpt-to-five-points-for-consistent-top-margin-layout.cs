// Title: Set worksheet top margin to 5 points (5/72 inch) using Aspose.Cells for .NET
// Common Searches: how to set worksheet top margin to five points with Aspose.Cells C# | Aspose.Cells C# example for page setup top margin in points | convert points to inches for Excel page margins using Aspose.Cells | C# code to apply consistent top margin across Excel worksheets with Aspose.Cells | Aspose.Cells page layout margin adjustment tutorial for .NET developers
// Tags: worksheet top margin adjustment Aspose.Cells | points to inches conversion for page margins | Aspose.Cells page setup API usage | C# Excel margin configuration | consistent top margin layout .NET

using Aspose.Cells;

// The example loads an Excel workbook, accesses the first worksheet, sets its top margin to five points (by converting 5 pts to inches with 5/72), and saves the workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook using the provided load rule
        Workbook workbook = LoadWorkbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Aspose.Cells uses inches for margins; 1 point = 1/72 inch.
        // Set the top margin to 5 points.
        sheet.PageSetup.TopMargin = 5.0 / 72.0;

        // Save the workbook using the provided save rule
        SaveWorkbook(workbook, "output.xlsx");
    }

    // Provided load rule placeholder
    static Workbook LoadWorkbook(string path)
    {
        // The actual implementation is supplied by the rule set.
        return new Workbook(path);
    }

    // Provided save rule placeholder
    static void SaveWorkbook(Workbook workbook, string path)
    {
        // The actual implementation is supplied by the rule set.
        workbook.Save(path);
    }
}
