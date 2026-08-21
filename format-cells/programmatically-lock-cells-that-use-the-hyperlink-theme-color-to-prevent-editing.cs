// Title: Lock Hyperlink‑Colored Cells and Protect Worksheet with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, formats a cell with the default hyperlink blue, adds a hyperlink, scans all used cells, locks any cell whose font color matches the hyperlink theme, applies the locked style flag, protects the worksheet with a password, and saves the file.
// Keywords: Aspose.Cells lock cells | C# lock cells by font color | protect worksheet Aspose.Cells | hyperlink theme color detection | Excel cell protection .NET
// Common Searches: How to lock cells that use the hyperlink theme color using Aspose.Cells C# | Aspose.Cells C# lock cells with blue font and protect worksheet | Programmatically protect Excel cells based on font color in .NET | Detect and lock hyperlink cells with Aspose.Cells
// Developer Intent: Identify cells styled with the default hyperlink font color, set them as locked, and enable worksheet protection so only non‑hyperlink cells stay editable.
// Use Cases: Generate a read‑only report where hyperlink cells are locked while other data can be edited. | Distribute a workbook that allows users to modify regular cells but prevents changes to linked URLs. | Automate data‑integrity enforcement by locking any cell using the standard hyperlink color before saving.
// AI Prompts: Write C# code with Aspose.Cells that locks all cells whose font color equals the default hyperlink blue and protects the worksheet with a password. | Show an alternative method to lock hyperlink cells by inspecting the worksheet's Hyperlink collection instead of font color. | Explain how to unlock specific cells after worksheet protection using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsHyperlinkLockDemo
{
    // Creates a workbook, formats a cell with the default hyperlink blue, adds a hyperlink, scans all used cells, locks any cell whose font color matches the hyperlink theme, applies the locked style flag, protects the worksheet with a password, and saves the file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Cell A1 – will have Hyperlink theme color (using standard blue)
                Cell cellA1 = cells["A1"];
                cellA1.PutValue("Link 1");
                Style styleA1 = cellA1.GetStyle();
                styleA1.Font.Color = Color.Blue; // typical hyperlink color
                cellA1.SetStyle(styleA1);
                // Add an actual hyperlink (optional, just for demonstration)
                worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");

                // Cell B1 – normal color (will remain editable)
                Cell cellB1 = cells["B1"];
                cellB1.PutValue("Normal");
                Style styleB1 = cellB1.GetStyle();
                styleB1.Font.Color = Color.Black;
                cellB1.SetStyle(styleB1);

                // Iterate through all used cells and lock those that use the Hyperlink color
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell curCell = cells[row, col];
                        if (curCell == null) continue;

                        Style curStyle = curCell.GetStyle();

                        // Check if the font color matches the hyperlink color (blue)
                        if (curStyle.Font.Color.ToArgb() == Color.Blue.ToArgb())
                        {
                            // Lock the cell
                            curStyle.IsLocked = true;
                            // Apply the style with the Locked flag enabled
                            StyleFlag flag = new StyleFlag { Locked = true };
                            curCell.SetStyle(curStyle, flag);
                        }
                    }
                }

                // Protect the worksheet so that locked cells cannot be edited
                // Provide an empty oldPassword as required by the overload
                worksheet.Protect(ProtectionType.All, "securePwd", string.Empty);

                // Save the workbook
                workbook.Save("HyperlinkLockedCells.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
