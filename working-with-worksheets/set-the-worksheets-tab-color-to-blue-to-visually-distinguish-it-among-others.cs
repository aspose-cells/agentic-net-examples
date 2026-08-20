// Title: How to Set an Excel Worksheet Tab Color to Blue with Aspose.Cells for .NET
// Description: Shows how to assign the TabColor property of a worksheet to blue using Aspose.Cells for .NET and save the workbook as an .xlsx file.
// Keywords: Aspose.Cells | C# worksheet tab color | Excel tab color programmatically | Worksheet.TabColor | set tab color blue | Aspose.Cells example | C# Excel styling | GitHub Aspose.Cells TabColor sample
// Common Searches: Aspose.Cells change worksheet tab color C# | set Excel sheet tab to blue using Aspose.Cells | Worksheet.TabColor property example | how to color Excel sheet tabs programmatically | C# code to set tab color in Excel file
// Developer Intent: Apply a blue color to a worksheet tab in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Highlight a summary sheet in a multi‑sheet report by giving it a blue tab. | Visually separate input, processing, and output worksheets with distinct tab colors, using blue for completed sections. | Create a status dashboard where blue tabs indicate tasks that have been finalized.
// AI Prompts: Generate a C# snippet that sets different TabColor values for several worksheets in the same workbook using Aspose.Cells. | Explain how to open an existing .xlsx file, locate a worksheet by name, and change its TabColor with Aspose.Cells. | Show how to define a custom RGB color and assign it to a worksheet tab instead of using predefined System.Drawing colors.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsTabColorDemo
{
    // Shows how to assign the TabColor property of a worksheet to blue using Aspose.Cells for .NET and save the workbook as an .xlsx file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the worksheet tab color to blue
            worksheet.TabColor = Color.Blue;

            // Save the workbook to a file
            workbook.Save("WorksheetWithBlueTab.xlsx", SaveFormat.Xlsx);
        }
    }
}
