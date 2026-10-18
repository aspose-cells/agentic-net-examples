// Title: How to apply a custom RGB fill color to header cells in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set RGB background color for header row | How to apply solid fill to a range of cells using Aspose.Cells .NET | Using Color.FromArgb to color Excel header cells with Aspose.Cells | Apply custom fill color to A1:D1 range in Aspose.Cells C# example | Aspose.Cells style flag all true to style header row
// Tags: apply solid fill color to range Aspose.Cells C# | set header row background RGB Aspose.Cells .NET | styleflag all true apply style Aspose.Cells | create range A1:D1 header styling Aspose.Cells | custom RGB background color Excel cells Aspose.Cells

using Aspose.Cells;
using System.Drawing;
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, defines the header range A1:D1, builds a style with a solid fill pattern, sets its foreground color to a light blue using Color.FromArgb(173,216,230), applies the style to the range with a StyleFlag (All = true), writes header text, and saves the file as HeaderColored.xlsx while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the header range (cells A1 to D1)
            AsposeRange headerRange = sheet.Cells.CreateRange("A1:D1");

            // Create a style for the header cells
            Style headerStyle = workbook.CreateStyle();

            // Set solid fill pattern
            headerStyle.Pattern = BackgroundType.Solid;

            // Apply custom RGB fill color (light blue)
            headerStyle.ForegroundColor = Color.FromArgb(173, 216, 230);

            // Apply the style to the header range
            StyleFlag flag = new StyleFlag { All = true };
            headerRange.ApplyStyle(headerStyle, flag);

            // Set header text
            sheet.Cells["A1"].PutValue("Header 1");
            sheet.Cells["B1"].PutValue("Header 2");
            sheet.Cells["C1"].PutValue("Header 3");
            sheet.Cells["D1"].PutValue("Header 4");

            // Save the workbook
            workbook.Save("HeaderColored.xlsx");
        }
        catch (Exception ex)
        {
            // Log or display the error
            System.Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
