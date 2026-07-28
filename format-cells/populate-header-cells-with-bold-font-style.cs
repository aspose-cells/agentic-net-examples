// Title: Apply Bold Font to Header Row Using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, writes headers to A1‑C1, defines a bold Style with a StyleFlag, applies it to row 0 via ApplyRowStyle, and saves as HeaderBold.xlsx.
// Keywords: Aspose.Cells C# bold header | ApplyRowStyle bold font | StyleFlag font bold Aspose | Excel header formatting .NET | C# set row style Aspose.Cells
// Common Searches: how to make header row bold Aspose.Cells C# | apply bold style to first row using Aspose.Cells | StyleFlag bold font example Aspose.Cells .NET | C# Aspose.Cells format header cells bold | Excel export with bold headers Aspose
// Developer Intent: Render the first worksheet row in bold without altering other cell attributes.
// Use Cases: Generating reports where column titles need visual emphasis. | Exporting data tables to Excel with a distinct header style. | Styling template rows programmatically while preserving existing cell formats.
// AI Prompts: Provide C# code that uses Aspose.Cells to bold the header row only. | Show how to create a StyleFlag for bold text and apply it with ApplyRowStyle in Aspose.Cells. | Explain combining bold font with other style attributes (e.g., background color) using StyleFlag in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, writes headers to A1‑C1, defines a bold Style with a StyleFlag, applies it to row 0 via ApplyRowStyle, and saves as HeaderBold.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate header cells
        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("Age");
        cells["C1"].PutValue("Country");

        // Create a style with bold font
        Style boldStyle = workbook.CreateStyle();
        boldStyle.Font.IsBold = true;

        // Define a StyleFlag to apply only the bold font attribute
        StyleFlag flag = new StyleFlag();
        flag.FontBold = true;

        // Apply the bold style to the first row (row index 0)
        sheet.Cells.ApplyRowStyle(0, boldStyle, flag);

        // Save the workbook
        workbook.Save("HeaderBold.xlsx");
    }
}
