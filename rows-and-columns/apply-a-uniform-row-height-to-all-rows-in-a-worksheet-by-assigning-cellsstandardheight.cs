using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – sets a uniform row height for all rows

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Apply a uniform row height (in points) to all rows in the worksheet
        worksheet.Cells.StandardHeight = 20; // adjust the value as needed

        // Save the workbook (lifecycle: save)
        workbook.Save("UniformRowHeight.xlsx");
    }
}