// Title: Apply a uniform row height to every row in an Excel worksheet with Aspose.Cells using Cells.StandardHeight (C#)
// Common Searches: C# Aspose.Cells set same row height for entire worksheet | how to use Cells.StandardHeight to define default row height in Excel file | Aspose.Cells C# change row height for all rows at once | set default row height for new workbook using Aspose.Cells | apply consistent row height across worksheet programmatically Aspose.Cells
// Tags: Aspose.Cells set worksheet row height | Cells.StandardHeight usage C# | default row height Excel file Aspose | uniform row height across rows Aspose.Cells | programmatic row height formatting Excel

using Aspose.Cells;
using System;

// The example creates a new Workbook, accesses its first Worksheet, sets Cells.StandardHeight to a specific value (e.g., 20 points) to give every row the same height, and saves the file as UniformRowHeight.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Apply a uniform row height (e.g., 20 points) to all rows
        sheet.Cells.StandardHeight = 20;

        // Save the workbook
        workbook.Save("UniformRowHeight.xlsx");
    }
}
