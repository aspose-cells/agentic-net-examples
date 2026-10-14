// Title: Replace <br> HTML tags with newline characters in all cells after importing an HTML workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an HTML file into an Aspose.Cells Workbook, iterates through every worksheet and cell, and substitutes any <br>, <br/>, or <br /> tags in string cells with '\n' before saving the workbook as XLSX. | Create a method in Aspose.Cells for .NET that post‑processes an imported HTML workbook by converting HTML line‑break tags to line‑feed characters in each cell's string value.
// Common Searches: Aspose.Cells .NET how to change <br> tags to line breaks after loading HTML | C# replace HTML line break tags in Excel cells using Aspose.Cells | convert <br/> to newline in cells after importing HTML with Aspose.Cells | post import HTML cleanup Aspose.Cells workbook C#
// Tags: aspocells html line break removal | aspocells cell value newline conversion | aspocells html import post processing | aspocells string cell cleanup

using Aspose.Cells;
using System;

// The program loads an HTML file into an Aspose.Cells Workbook, scans every worksheet and cell, replaces any <br>, <br/>, or <br /> tags found in string cells with line feed characters, updates the cell values, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        // Load the HTML file into a workbook
        Workbook workbook = new Workbook("input.html");

        // Iterate through all worksheets and cells
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue;
                        // Replace <br>, <br/>, and <br /> with line feed characters
                        string replaced = original
                            .Replace("<br>", "\n")
                            .Replace("<br/>", "\n")
                            .Replace("<br />", "\n");

                        if (replaced != original)
                        {
                            cell.PutValue(replaced);
                        }
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
