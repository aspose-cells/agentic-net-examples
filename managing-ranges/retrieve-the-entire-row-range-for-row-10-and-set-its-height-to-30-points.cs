// Title: Retrieve row 10 range and set its height to 30 points with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# retrieve entire row 10 range | Set row height to 30 points using Aspose.Cells .NET | How to change row 10 height programmatically in Aspose.Cells | Rows collection height property example Aspose.Cells C# | Adjust row dimensions in an Excel workbook with Aspose.Cells
// Tags: row height property Aspose.Cells | retrieve row range Aspose.Cells | rows collection height .NET | adjust row height programmatically Aspose.Cells | excel row height manipulation C#

using Aspose.Cells;
using System;

// The example creates a new Workbook, accesses the first worksheet, retrieves row 10 (zero‑based index 9) via the Rows collection, sets its Height property to 30 points, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Retrieve the entire row 10 (zero‑based index 9)
        Row row10 = sheet.Cells.Rows[9];

        // Set the row height to 30 points
        row10.Height = 30;

        // Save the workbook (lifecycle: save)
        workbook.Save("Output.xlsx");
    }
}
