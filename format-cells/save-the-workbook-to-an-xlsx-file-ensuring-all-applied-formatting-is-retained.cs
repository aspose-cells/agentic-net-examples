// Title: Save a formatted workbook to XLSX while preserving all cell styles using Aspose.Cells for .NET
// Common Searches: c# aspso.cells save workbook to xlsx with header style retained | how to preserve cell formatting when exporting to xlsx using Aspose.Cells | auto‑fit columns then save Excel file with Aspose.Cells .NET | export a styled worksheet to XLSX format with Aspose.Cells
// Tags: save workbook to xlsx preserving formatting aspso.cells | apply bold header style aspso.cells | auto fit columns before saving aspso.cells | export styled worksheet to xlsx aspso.cells | retain cell styles during export aspso.cells

using Aspose.Cells;
using System.Drawing;

// The example creates a workbook, adds data, applies a bold gray header style, auto‑fits columns, and saves the file as FormattedReport.xlsx in XLSX format, ensuring all formatting is retained.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate some data
        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("Score");
        cells["A2"].PutValue("Alice");
        cells["B2"].PutValue(85);
        cells["A3"].PutValue("Bob");
        cells["B3"].PutValue(92);

        // Create a style for the header row
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.ForegroundColor = Color.LightGray;
        headerStyle.Pattern = BackgroundType.Solid;

        // Apply the style to header cells
        StyleFlag flag = new StyleFlag();
        flag.All = true;
        cells["A1"].SetStyle(headerStyle, flag);
        cells["B1"].SetStyle(headerStyle, flag);

        // Auto-fit columns to display content properly
        sheet.AutoFitColumns();

        // Save the workbook to an XLSX file, preserving all formatting
        workbook.Save("FormattedReport.xlsx", SaveFormat.Xlsx);
    }
}
