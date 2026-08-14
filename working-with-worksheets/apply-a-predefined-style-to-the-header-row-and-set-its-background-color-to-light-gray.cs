// Title: Style Header Row with Light Gray Background Using Aspose.Cells for .NET
// Description: Creates a workbook, defines a solid LightGray style with bold centered font, applies it to the first row via ApplyRowStyle and a full StyleFlag, then saves as HeaderRowStyled.xlsx.
// Keywords: Aspose.Cells C# header style | ApplyRowStyle | StyleFlag All | light gray header background | Excel header formatting .NET | solid fill Aspose.Cells | centered bold font | worksheet header row style
// Common Searches: Aspose.Cells set header row background color | C# apply style to first row Excel | How to use StyleFlag with ApplyRowStyle | Create gray header row Aspose.Cells | Bold centered header in Excel using Aspose
// Developer Intent: Apply a predefined style with a light‑gray background and bold centered text to the worksheet’s header row.
// Use Cases: Automated report generation where column headings are highlighted for quick visual scanning. | Ensuring consistent header formatting across multiple exported spreadsheets in a data‑pipeline. | Producing printable tables with a distinct header row to improve readability.
// AI Prompts: Generate C# code using Aspose.Cells to create a Style with solid LightGray fill, bold black 12‑pt font, centered alignment, and apply it to row 0 via ApplyRowStyle. | Show how to define a reusable Style and a StyleFlag set to All=true for formatting header rows in several worksheets. | Demonstrate applying the same header style to rows 0‑2 using a loop and ApplyRowStyle in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHeaderStyleDemo
{
    // Creates a workbook, defines a solid LightGray style with bold centered font, applies it to the first row via ApplyRowStyle and a full StyleFlag, then saves as HeaderRowStyled.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate header cells (optional, just for demonstration)
            cells[0, 0].PutValue("ID");
            cells[0, 1].PutValue("Name");
            cells[0, 2].PutValue("Date");

            // Create a style for the header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Pattern = BackgroundType.Solid;               // Use solid fill
            headerStyle.ForegroundColor = Color.LightGray;            // Light gray background
            headerStyle.Font.IsBold = true;                           // Make text bold
            headerStyle.Font.Color = Color.Black;                     // Font color
            headerStyle.Font.Size = 12;                               // Font size
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;
            headerStyle.VerticalAlignment = TextAlignmentType.Center;

            // Define a style flag to apply all formatting properties
            StyleFlag flag = new StyleFlag { All = true };

            // Apply the style to the first (header) row (row index 0)
            cells.ApplyRowStyle(0, headerStyle, flag);

            // Save the workbook
            workbook.Save("HeaderRowStyled.xlsx");
        }
    }
}
