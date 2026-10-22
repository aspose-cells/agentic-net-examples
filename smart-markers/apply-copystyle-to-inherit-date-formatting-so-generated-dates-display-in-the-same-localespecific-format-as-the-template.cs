// Title: Copy a template cell's locale‑specific date format to generated date cells with Aspose.Cells for .NET
// AI Prompts: Retrieve the style from a template date cell and apply it to a range of new cells while inserting DateTime values using Aspose.Cells. | Use Aspose.Cells to inherit the locale‑specific date format from cell A1 and set it on dynamically created date cells in C#. | Programmatically apply a source cell's number format to multiple cells in a workbook with Aspose.Cells.
// Common Searches: Aspose.Cells copy date format from one cell to another in C# | How to preserve locale specific date formatting when writing dates with Aspose.Cells .NET | Apply template cell style to generated cells using Aspose.Cells workbook | Set number format for multiple cells based on a source cell in Aspose.Cells | C# Aspose.Cells inherit date style from template worksheet
// Tags: copy cell style Aspose.Cells C# | inherit locale date format Aspose.Cells | apply template date style to generated cells | set number format from source cell Aspose.Cells | Aspose.Cells date formatting example

using Aspose.Cells;
using System;

// The example loads a template workbook, extracts the style (including locale‑specific date format) from cell A1, writes three DateTime values into column B, applies the extracted style to each new cell, and saves the result as Result.xlsx.
class Program
{
    static void Main()
    {
        // Load the template workbook (using the provided load rule)
        Workbook workbook = new Workbook("Template.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Assume cell A1 in the template contains the desired date format
        Cell templateDateCell = sheet.Cells["A1"];
        Style dateStyle = templateDateCell.GetStyle();

        // Sample dates to write
        DateTime[] dates = new DateTime[]
        {
            DateTime.Now,
            DateTime.Now.AddDays(1),
            DateTime.Now.AddDays(2)
        };

        // Write dates to column B and copy the date style from the template
        for (int i = 0; i < dates.Length; i++)
        {
            // Row i, column 1 corresponds to B{i+1}
            Cell newCell = sheet.Cells[i, 1];
            newCell.PutValue(dates[i]);

            // Apply the same style (including locale‑specific date format)
            newCell.SetStyle(dateStyle);
        }

        // Save the workbook (using the provided save rule)
        workbook.Save("Result.xlsx");
    }
}
