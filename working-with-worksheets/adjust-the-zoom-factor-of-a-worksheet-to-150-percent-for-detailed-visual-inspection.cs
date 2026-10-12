// Title: Set worksheet zoom level to 150% with Aspose.Cells for .NET (C# example)
// AI Prompts: Write C# code that opens a workbook with Aspose.Cells, sets the first worksheet's Zoom property to 150, and saves the file. | Demonstrate how to programmatically change the view scale of a specific worksheet to 150 percent using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# how to change worksheet zoom to 150 percent | programmatically set Excel sheet zoom level using Aspose.Cells .NET | example code for adjusting worksheet view scale with Aspose.Cells | increase worksheet zoom factor in a generated workbook using Aspose.Cells C#
// Tags: worksheet zoom Aspose.Cells C# | adjust view scale Excel Aspose.Cells | set worksheet zoom property .NET | 150 percent zoom Aspose.Cells example

using System;
using Aspose.Cells;

// // Creates (or loads) a workbook, accesses the first worksheet, sets its Zoom property to 150%, and saves the result as ZoomedWorksheet.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the zoom factor to 150%
        sheet.Zoom = 150;

        // Save the workbook to a file
        workbook.Save("ZoomedWorksheet.xlsx");
    }
}
