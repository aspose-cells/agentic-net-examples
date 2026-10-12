// Title: Replace EM DASH and smart quotes with a custom font in Excel cells before saving as PDF using Aspose.Cells for .NET
// AI Prompts: Search the workbook for cells containing EM DASH (U+2014) or smart quotes (U+201C/U+201D), set their Font.Name to MyCustomFont, and generate a PDF file. | Iterate over all worksheets in a .xlsx, detect any of the specified Unicode characters, apply a custom font style to those cells, then export the workbook to PDF with Aspose.Cells.
// Common Searches: Aspose.Cells C# replace EM DASH with custom font before PDF conversion | How to change font for cells that contain smart quotes in an Excel file using Aspose.Cells | Detect specific Unicode characters and apply a different font when exporting to PDF with Aspose.Cells | C# iterate through Excel cells and set custom font for Unicode characters for PDF output | Aspose.Cells selective font styling for Unicode characters during PDF save
// Tags: unicode character font replacement Aspose.Cells | apply custom font to Excel cells Aspose.Cells | selective font styling during PDF export Aspose.Cells | detect EM DASH smart quotes Aspose.Cells | excel to pdf custom font mapping Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an .xlsx file, scans each cell for EM DASH and smart‑quote characters, assigns a custom font (MyCustomFont) to those cells, and then saves the workbook as a PDF using Aspose.Cells.
class ReplaceUnicodeAndSavePdf
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Define the Unicode characters that need a custom font
        // Example: replace EM DASH (U+2014) and LEFT DOUBLE QUOTATION MARK (U+201C)
        var targetChars = new HashSet<char> { '\u2014', '\u201C', '\u201D' };

        // Define the custom font name to apply
        const string customFontName = "MyCustomFont";

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the used range of the worksheet
            var cells = sheet.Cells;
            var maxRow = cells.MaxDataRow;
            var maxCol = cells.MaxDataColumn;

            // Loop through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only string cells
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue;
                        bool containsTarget = false;

                        // Check if the cell contains any of the target Unicode characters
                        foreach (char ch in text)
                        {
                            if (targetChars.Contains(ch))
                            {
                                containsTarget = true;
                                break;
                            }
                        }

                        if (containsTarget)
                        {
                            // Apply the custom font to the entire cell
                            Style style = cell.GetStyle();
                            style.Font.Name = customFontName;
                            cell.SetStyle(style);
                        }
                    }
                }
            }
        }

        // Save the workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
