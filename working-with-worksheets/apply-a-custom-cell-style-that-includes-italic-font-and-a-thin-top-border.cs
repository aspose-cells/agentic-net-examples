// Title: C# – Apply Italic Font and Thin Top Border to a Cell with Aspose.Cells
// Description: This Aspose.Cells for .NET example creates a workbook, writes text to cell B2, builds a custom style that makes the font italic and adds a thin black top border, uses a StyleFlag to apply only those attributes, and saves the file as an XLSX document.
// Keywords: Aspose.Cells | C# | .NET | custom cell style | italic font | top border | thin border | StyleFlag | cell formatting | Excel export | Workbook
// Common Searches: Aspose.Cells set italic font C# | How to add top border to a cell Aspose.Cells | StyleFlag example Aspose.Cells .NET | Create custom style with border Aspose.Cells | Apply thin border to specific cell C#
// Developer Intent: Developer wants to format a single cell with italic text and a thin top border using a custom style in Aspose.Cells for .NET.
// Use Cases: Design report headers where only the top edge is highlighted and text is italic. | Separate summary rows in financial sheets with a thin top line while keeping the font slanted. | Build reusable templates that apply italic headings with a top border across multiple worksheets.
// AI Prompts: Write C# code that applies bold font and double bottom border to a range using Aspose.Cells. | Show how to create a reusable style with multiple StyleFlag options and apply it to several worksheets. | Demonstrate setting different border styles on each side of a cell while preserving existing font settings in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsCustomStyleDemo
{
    // This Aspose.Cells for .NET example creates a workbook, writes text to cell B2, builds a custom style that makes the font italic and adds a thin black top border, uses a StyleFlag to apply only those attributes, and saves the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Target cell (e.g., B2)
            Cell cell = cells["B2"];
            cell.PutValue("Italic with Top Border");

            // Create a custom style
            Style customStyle = workbook.CreateStyle();

            // Set italic font
            customStyle.Font.IsItalic = true;

            // Configure a thin top border
            customStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            customStyle.Borders[BorderType.TopBorder].Color = Color.Black;

            // Define which style elements to apply using StyleFlag
            StyleFlag flag = new StyleFlag
            {
                FontItalic = true,   // Apply italic font
                TopBorder = true     // Apply top border
            };

            // Apply the style to the cell with the specified flags
            cell.SetStyle(customStyle, flag);

            // Save the workbook
            workbook.Save("CustomStyleDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
