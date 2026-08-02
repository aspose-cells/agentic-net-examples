// Title: C# – Apply Light Gray Background Style to Worksheet Header Row with Aspose.Cells
// Description: Creates a new workbook, defines a solid light‑gray background style, uses a StyleFlag to apply all style attributes, styles the first row (header) of the worksheet, and saves the file as HeaderStyle.xlsx.
// Keywords: Aspose.Cells header row style C# | set row background color Aspose.Cells | StyleFlag apply style Aspose.Cells | solid light gray background Aspose.Cells | C# Excel header formatting Aspose
// Common Searches: Aspose.Cells set header row background color C# | How to apply a style to the first row in Aspose.Cells | C# Aspose.Cells StyleFlag example | Create gray header row in Excel with Aspose.Cells | Apply solid background to worksheet row using Aspose.Cells
// Developer Intent: Apply a predefined style with a light gray background to the worksheet’s header row.
// Use Cases: Standardize header appearance for automated Excel reports. | Highlight column titles to improve spreadsheet readability. | Reuse a single style across multiple worksheets in a workbook.
// AI Prompts: Generate C# Aspose.Cells code that adds bold font, center alignment, and a light gray background to a worksheet header row. | Show how to define a reusable style and apply it to several rows and columns with StyleFlag in Aspose.Cells C#. | Explain how to modify the header style to include custom font size while keeping the solid light gray background.

using System.Drawing;
using Aspose.Cells;

// Creates a new workbook, defines a solid light‑gray background style, uses a StyleFlag to apply all style attributes, styles the first row (header) of the worksheet, and saves the file as HeaderStyle.xlsx.
class ApplyHeaderStyle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a style with a solid light gray background
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Pattern = BackgroundType.Solid;
        headerStyle.BackgroundColor = Color.LightGray;

        // Define a style flag to apply all style attributes
        StyleFlag flag = new StyleFlag { All = true };

        // Apply the style to the first row (header row)
        worksheet.Cells.Rows[0].ApplyStyle(headerStyle, flag);

        // Save the workbook
        workbook.Save("HeaderStyle.xlsx");
    }
}
