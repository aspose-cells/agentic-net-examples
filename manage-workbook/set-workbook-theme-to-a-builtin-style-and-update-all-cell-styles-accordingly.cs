// Title: Apply a Built‑in Workbook Theme to All Cells with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, set a built‑in style (e.g., Good) as Workbook.DefaultStyle, iterate through every worksheet and populated cell, apply the default style, and save the result as an Excel file using Aspose.Cells for C#.
// Keywords: Aspose.Cells C# | set built‑in theme | BuiltinStyleType | Workbook.DefaultStyle | apply style to all cells | Excel theme Aspose | cell formatting C# | Aspose.Cells API | default workbook style | iterate cells Aspose
// Common Searches: How to set a built‑in theme in Aspose.Cells C# | Apply default workbook style to every cell using Aspose.Cells | Change Excel workbook theme to Good style with Aspose.Cells | C# Aspose.Cells update all cell styles after changing theme | Set Workbook.DefaultStyle and refresh formatting in all worksheets
// Developer Intent: Assign a built‑in theme to a workbook and propagate the style to every existing cell.
// Use Cases: Create a new report where the Good built‑in style is applied uniformly to all cells after data entry. | Retheme an existing spreadsheet by switching Workbook.DefaultStyle to another BuiltinStyleType and re‑applying it across all worksheets. | Automate consistent branding for generated Excel files by enforcing a single default style on every cell programmatically.
// AI Prompts: Generate C# code with Aspose.Cells that changes the workbook theme to the 'Bad' built‑in style and updates all cells in every worksheet. | Write a reusable method for Aspose.Cells that accepts a BuiltinStyleType parameter, sets Workbook.DefaultStyle, and reapplies the style to all cells in a given workbook. | Provide a step‑by‑step guide for applying a custom theme to an existing Excel file using Aspose.Cells, including style propagation and saving the file.

using System;
using Aspose.Cells;

namespace AsposeCellsThemeDemo
{
    // Demonstrates how to create a workbook, set a built‑in style (e.g., Good) as Workbook.DefaultStyle, iterate through every worksheet and populated cell, apply the default style, and save the result as an Excel file using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text 1");
            sheet.Cells["B2"].PutValue("Sample Text 2");
            sheet.Cells["C3"].PutValue("Sample Text 3");

            // Create a built‑in style (e.g., Good) and set it as the default style
            Style builtinStyle = workbook.CreateBuiltinStyle(BuiltinStyleType.Good);
            workbook.DefaultStyle = builtinStyle;

            // Apply the default style to all existing cells in all worksheets
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Cells cells = ws.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        // Apply the default (built‑in) style to each cell
                        cell.SetStyle(workbook.DefaultStyle);
                    }
                }
            }

            // Save the workbook
            workbook.Save("BuiltInThemeDemo.xlsx");
        }
    }
}
