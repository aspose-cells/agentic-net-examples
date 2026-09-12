// Title: Generate an Excel worksheet that displays a preview grid of all workbook theme colors with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to list every ThemeColorType index in column A and fill column B with a solid background of the corresponding theme color. | Extend the program to include the theme color name beside each preview cell and save the workbook as both .xlsx and .pdf files. | Create a reusable method that accepts an array of Color objects and returns a styled range that shows each color as a centered cell background.
// Common Searches: how to list all ThemeColorType values and show their colors in an Excel sheet with Aspose.Cells C# | Aspose.Cells C# generate color swatch grid for workbook theme colors | export theme color preview worksheet to PDF using Aspose.Cells .NET | display Excel theme colors as solid cell backgrounds programmatically
// Tags: Aspose.Cells generate theme color preview worksheet | C# create Excel color swatch grid | retrieve workbook ThemeColorType values Aspose | apply solid background style to cells Aspose.Cells | export theme color preview to PDF .NET

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, extracts every ThemeColorType defined in the workbook, writes the index in column A, applies each theme color as a solid background to column B cells, adds a bold header row, and saves the file as ThemeColorPreview.xlsx.
class ThemeColorPreview
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve all theme colors defined in the workbook
            ThemeColorType[] themeEnums = (ThemeColorType[])Enum.GetValues(typeof(ThemeColorType));
            Color[] themeColors = new Color[themeEnums.Length];
            for (int i = 0; i < themeEnums.Length; i++)
            {
                themeColors[i] = workbook.GetThemeColor(themeEnums[i]);
            }

            // Set column widths for better visibility
            sheet.Cells.SetColumnWidth(0, 20); // Column for color index
            sheet.Cells.SetColumnWidth(1, 15); // Column for color preview

            // Add a header row
            sheet.Cells["A1"].PutValue("Theme Color Index");
            sheet.Cells["B1"].PutValue("Preview");

            // Apply bold style to header
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;
            headerStyle.VerticalAlignment = TextAlignmentType.Center;
            StyleFlag headerFlag = new StyleFlag
            {
                Font = true,
                HorizontalAlignment = true,
                VerticalAlignment = true
            };
            sheet.Cells["A1:B1"].SetStyle(headerStyle, headerFlag);

            // Populate the grid with each theme color
            for (int i = 0; i < themeColors.Length; i++)
            {
                int row = i + 2; // Start from row 2 (after header)

                // Write the theme color index
                sheet.Cells[row, 0].PutValue(i);

                // Create a style with the theme color as background
                Style colorStyle = workbook.CreateStyle();
                colorStyle.ForegroundColor = themeColors[i];
                colorStyle.Pattern = BackgroundType.Solid;
                colorStyle.HorizontalAlignment = TextAlignmentType.Center;
                colorStyle.VerticalAlignment = TextAlignmentType.Center;

                // Apply the style to the preview cell
                StyleFlag flag = new StyleFlag
                {
                    CellShading = true,
                    HorizontalAlignment = true,
                    VerticalAlignment = true
                };
                sheet.Cells[row, 1].SetStyle(colorStyle, flag);
            }

            // Save the workbook to a file
            workbook.Save("ThemeColorPreview.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
