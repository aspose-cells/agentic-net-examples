// Title: Set Worksheet Tab Color to Teal with Aspose.Cells for .NET (C#)
// Description: A concise guide showing how to change an Excel worksheet's tab color to teal using Aspose.Cells for .NET. The sample creates a workbook, selects the first sheet, assigns Color.Teal to the TabColor property, and saves the file.
// Keywords: Aspose.Cells worksheet tab color | C# set Excel tab color | TabColor property .NET | change sheet tab color programmatically | teal tab color Aspose.Cells | Excel workbook styling C#
// Common Searches: how to set worksheet tab color using Aspose.Cells C# | Aspose.Cells change Excel sheet tab to teal | C# code for setting tab color in Excel workbook | programmatically modify worksheet tab color .NET | Aspose.Cells TabColor example
// Developer Intent: Apply a teal color to a worksheet tab via Aspose.Cells.
// Use Cases: Highlight a summary or dashboard sheet with a distinct teal tab. | Apply corporate branding colors to specific tabs in automated reports. | Mark newly generated worksheets for easy identification during runtime.
// AI Prompts: Generate C# code that uses Aspose.Cells to set a worksheet's TabColor to teal and save the workbook. | Show how to assign different TabColor values to multiple worksheets in a single workbook with Aspose.Cells. | Explain the steps to read, modify, and persist the TabColor property of an Excel sheet using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;

// A concise guide showing how to change an Excel worksheet's tab color to teal using Aspose.Cells for .NET. The sample creates a workbook, selects the first sheet, assigns Color.Teal to the TabColor property, and saves the file.
class SetWorksheetTabColor
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Access the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the worksheet tab color to teal
        worksheet.TabColor = Color.Teal;

        // Save the workbook to a file
        workbook.Save("WorksheetWithTealTab.xlsx");
    }
}
