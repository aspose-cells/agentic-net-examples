// Title: How to apply a custom theme border color to a cell range in the second worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set custom border color for a range on a specific worksheet | Apply thin orange borders to B2:D5 in the second sheet with Aspose.Cells | Use StyleFlag to format only borders of a range in Aspose.Cells .NET | Create and apply a custom theme color to cell borders in Aspose.Cells workbook | Add second worksheet and style its range borders with a custom color in C#
// Tags: custom theme border styling Aspose.Cells | range border formatting second worksheet | StyleFlag border-only application Aspose.Cells | thin orange border style .NET | Aspose.Cells workbook border color customization

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, adds a second worksheet named "SecondSheet", defines the range B2:D5, builds a style with thin orange custom borders on all sides, uses a StyleFlag to apply only the border formatting, applies the style to the range, and saves the file as CustomThemeBorder.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (first worksheet is created by default)
            Workbook workbook = new Workbook();

            // Add a second worksheet and give it a name
            int secondSheetIndex = workbook.Worksheets.Add();
            Worksheet secondSheet = workbook.Worksheets[secondSheetIndex];
            secondSheet.Name = "SecondSheet";

            // Define the range whose borders will be styled (e.g., B2:D5)
            AsposeRange targetRange = secondSheet.Cells.CreateRange("B2", "D5");

            // Create a style object to hold border settings
            Style borderStyle = workbook.CreateStyle();

            // Define a custom color (you can replace the ARGB values with any theme color you need)
            Color customColor = Color.FromArgb(255, 128, 0); // Orange-like custom color

            // Apply thin borders on all sides using the custom color
            borderStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, customColor);
            borderStyle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, customColor);
            borderStyle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, customColor);
            borderStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, customColor);

            // Specify that only border formatting should be applied
            StyleFlag styleFlag = new StyleFlag { Borders = true };

            // Apply the style to the defined range
            targetRange.ApplyStyle(borderStyle, styleFlag);

            // Save the workbook to a file
            string outputPath = "CustomThemeBorder.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
