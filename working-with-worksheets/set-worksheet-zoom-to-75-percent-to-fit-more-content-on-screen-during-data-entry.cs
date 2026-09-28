// Title: How to set a worksheet’s zoom level to 75 % using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new workbook, accesses the first worksheet, sets its Zoom property to 75, and saves the file. | Show how to adjust the view scale of a specific worksheet to 75 % before saving with Aspose.Cells. | Demonstrate changing the worksheet zoom factor programmatically in C# using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# set worksheet zoom to 75 percent before saving | programmatically change Excel sheet zoom level using Aspose.Cells .NET | C# example for adjusting worksheet view scale with Aspose.Cells
// Tags: Aspose.Cells set worksheet zoom | C# worksheet view scale Aspose | adjust Excel sheet zoom programmatically | Aspose.Cells workbook save with custom zoom | worksheet Zoom property .NET

using Aspose.Cells;
using System;

// // Creates a new workbook, sets the first worksheet's Zoom property to 75 %, and saves as ZoomedWorksheet.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the zoom factor to 75%
        sheet.Zoom = 75;

        // Save the workbook to a file
        workbook.Save("ZoomedWorksheet.xlsx");
    }
}
