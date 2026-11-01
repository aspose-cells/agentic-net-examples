// Title: Set narrow 0.25‑inch margins and landscape orientation for every worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells set 0.25 inch margins for all worksheets C# | C# Aspose.Cells change workbook page orientation to landscape | Apply identical page setup to each worksheet in Aspose.Cells .NET | Programmatically set narrow top and bottom margins in Aspose.Cells | Configure page layout margins and orientation for an Aspose.Cells workbook
// Tags: worksheet margin configuration Aspose.Cells | landscape orientation page setup Aspose.Cells | apply page setup to all worksheets .NET | narrow worksheet margins Aspose.Cells | C# workbook print settings Aspose.Cells

using Aspose.Cells;

// The example creates a workbook, iterates through all worksheets, sets every margin (top, bottom, left, right, header, footer) to 18 points (0.25 inch), changes the page orientation to landscape, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        var workbook = new Workbook();

        // Apply page setup to each worksheet in the workbook
        foreach (Worksheet ws in workbook.Worksheets)
        {
            PageSetup pageSetup = ws.PageSetup;

            // Set narrow margins (0.25 inch = 18 points)
            pageSetup.TopMargin = 18;      // points
            pageSetup.BottomMargin = 18;   // points
            pageSetup.LeftMargin = 18;     // points
            pageSetup.RightMargin = 18;    // points
            pageSetup.HeaderMargin = 18;   // points
            pageSetup.FooterMargin = 18;   // points

            // Set orientation to landscape
            pageSetup.Orientation = PageOrientationType.Landscape;
        }

        // Save the workbook with the new page settings
        workbook.Save("output.xlsx");
    }
}
