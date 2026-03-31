using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook from disk
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath); // uses the Workbook(string) constructor

        // Get the first worksheet in the workbook
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Create a style for the header row (A1:D1)
        // -------------------------------------------------
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Color = Color.White;
        headerStyle.ForegroundColor = Color.DarkBlue;
        headerStyle.Pattern = BackgroundType.Solid;
        headerStyle.HorizontalAlignment = TextAlignmentType.Center;

        // Apply the header style to the range A1:D1
        StyleFlag headerFlag = new StyleFlag();
        headerFlag.All = true; // apply all style attributes
        sheet.Cells.CreateRange("A1:D1").ApplyStyle(headerStyle, headerFlag);

        // -------------------------------------------------
        // Create a style for data rows (A2:D10)
        // -------------------------------------------------
        Style dataStyle = workbook.CreateStyle();
        dataStyle.Font.Color = Color.Black;
        dataStyle.ForegroundColor = Color.LightYellow;
        dataStyle.Pattern = BackgroundType.Solid;
        dataStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
        dataStyle.Borders[BorderType.BottomBorder].Color = Color.Gray;

        StyleFlag dataFlag = new StyleFlag();
        dataFlag.All = true;
        sheet.Cells.CreateRange("A2:D10").ApplyStyle(dataStyle, dataFlag);

        // -------------------------------------------------
        // Adjust column widths and row height
        // -------------------------------------------------
        sheet.Cells.SetColumnWidth(0, 20); // Column A
        sheet.Cells.SetColumnWidth(1, 30); // Column B
        sheet.Cells.SetColumnWidth(2, 15); // Column C
        sheet.Cells.SetColumnWidth(3, 25); // Column D

        sheet.Cells.SetRowHeight(0, 25); // Header row height

        // -------------------------------------------------
        // Additional worksheet formatting
        // -------------------------------------------------
        sheet.IsGridlinesVisible = false;      // hide gridlines
        sheet.TabColor = Color.Green;          // set worksheet tab color

        // -------------------------------------------------
        // Save the modified workbook to a new file
        // -------------------------------------------------
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx); // uses Save(string, SaveFormat)
    }
}