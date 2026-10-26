// Title: Show worksheet gridlines on screen while preventing them from printing using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# display gridlines in Excel viewer but not in printed output | How to keep gridlines visible on screen and hide them when printing with Aspose.Cells | Set Worksheet.IsGridlinesVisible true and PageSetup.PrintGridlines false in C# | Disable printed gridlines while retaining on‑screen gridlines in an Aspose.Cells workbook | Configure screen vs print gridline visibility for Excel files using Aspose.Cells .NET
// Tags: Worksheet.IsGridlinesVisible property C# | PageSetup.PrintGridlines setting Aspose.Cells | gridline visibility control Excel .NET | suppress printed gridlines Aspose.Cells | display gridlines in Excel viewer Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, enables on‑screen gridlines with Worksheet.IsGridlinesVisible, disables gridlines in printed output via PageSetup.PrintGridlines, adds sample data, and saves the file as GridlinesDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook.
        Workbook workbook = new Workbook();

        // Access the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];

        // Enable gridlines for on‑screen display.
        // This makes gridlines visible when the workbook is opened in the viewer.
        sheet.IsGridlinesVisible = true;

        // Hide gridlines in the printed output.
        // The PrintGridlines property controls whether gridlines are printed.
        sheet.PageSetup.PrintGridlines = false;

        // (Optional) Add some sample data to visualize the gridlines.
        sheet.Cells["A1"].PutValue("Header 1");
        sheet.Cells["B1"].PutValue("Header 2");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["B2"].PutValue(456);

        // Save the workbook to a file.
        workbook.Save("GridlinesDemo.xlsx", SaveFormat.Xlsx);
    }
}
