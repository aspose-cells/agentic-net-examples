// Title: How to set the height of row 10 to 30 points with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set specific row height in Excel | change height of row 10 using Aspose.Cells API | programmatically adjust row height to 30 points in Aspose.Cells | retrieve a row range and modify its height with Aspose.Cells C# | set Excel row height via Aspose.Cells .NET library
// Tags: Aspose.Cells set row height | C# Aspose.Cells modify row height | Excel row height adjustment Aspose.Cells | Aspose.Cells retrieve row range C# | row height 30 points Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new Workbook, accesses the first worksheet, obtains row 10 (zero‑based index 9) via the Rows collection, sets its Height property to 30 points, and saves the file as RowHeightExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Retrieve the entire row range for row 10 (zero‑based index 9)
        Row row = sheet.Cells.Rows[9];

        // Set the row height to 30 points
        row.Height = 30;

        // Save the workbook
        workbook.Save("RowHeightExample.xlsx");
    }
}
