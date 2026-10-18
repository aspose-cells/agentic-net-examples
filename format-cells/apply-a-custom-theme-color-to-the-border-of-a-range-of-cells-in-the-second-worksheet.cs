// Title: Set a custom theme color for the borders of a cell range on the second worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set custom theme color for range borders on second worksheet | How to apply thin blue borders to B2:D5 in a specific worksheet with Aspose.Cells | Using StyleFlag to modify only border colors in an Aspose.Cells workbook | Apply custom RGB border color to a cell range in Aspose.Cells .NET | Create and style borders for a range on a newly added worksheet in C#
// Tags: custom theme border color Aspose.Cells | apply style to range borders C# | second worksheet border styling Aspose.Cells | StyleFlag borders only .NET | RGB border color Aspose.Cells

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Add a second worksheet (index 1)
Aspose.Cells.Worksheet sheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
sheet2.Name = "SecondSheet";

// Define the range to which the border will be applied (e.g., B2:D5)
Aspose.Cells.Range range = sheet2.Cells.CreateRange("B2", "D5");

// Create a style object
Aspose.Cells.Style style = workbook.CreateStyle();

// Set the border style (continuous line)
style.Borders[Aspose.Cells.BorderType.BottomBorder].LineStyle = Aspose.Cells.CellBorderType.Thin;
style.Borders[Aspose.Cells.BorderType.TopBorder].LineStyle = Aspose.Cells.CellBorderType.Thin;
style.Borders[Aspose.Cells.BorderType.LeftBorder].LineStyle = Aspose.Cells.CellBorderType.Thin;
style.Borders[Aspose.Cells.BorderType.RightBorder].LineStyle = Aspose.Cells.CellBorderType.Thin;

// Apply a custom theme color to the borders (e.g., a shade of blue)
System.Drawing.Color customThemeColor = System.Drawing.Color.FromArgb(0, 112, 192); // Custom theme color
style.Borders[Aspose.Cells.BorderType.BottomBorder].Color = customThemeColor;
style.Borders[Aspose.Cells.BorderType.TopBorder].Color = customThemeColor;
style.Borders[Aspose.Cells.BorderType.LeftBorder].Color = customThemeColor;
style.Borders[Aspose.Cells.BorderType.RightBorder].Color = customThemeColor;

// Apply the style to the entire range
range.ApplyStyle(style, new Aspose.Cells.StyleFlag()
{
    // Apply border settings only
    Borders = true
});

// Save the workbook to a file
workbook.Save("Output.xlsx");
