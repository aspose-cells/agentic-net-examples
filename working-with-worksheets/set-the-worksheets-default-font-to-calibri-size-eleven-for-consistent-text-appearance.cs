using System;
using Aspose.Cells;

class SetWorksheetDefaultFont
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the workbook's default style settings to Calibri, size 11
        workbook.Settings.DefaultStyleSettings.FontName = "Calibri";
        workbook.Settings.DefaultStyleSettings.FontSize = 11.0;

        // Also update the DefaultStyle object (used when new cells are created)
        Style defaultStyle = workbook.DefaultStyle;
        defaultStyle.Font.Name = "Calibri";
        defaultStyle.Font.Size = 11;
        workbook.DefaultStyle = defaultStyle;

        // Apply the default style to the first worksheet to affect existing cells
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells.ApplyStyle(workbook.CreateStyle(), new StyleFlag() { All = true });

        // Save the workbook
        workbook.Save("WorksheetWithDefaultFont.xlsx");
    }
}