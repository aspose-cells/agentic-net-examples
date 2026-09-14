// Title: Set a teal tab color for a specific worksheet in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to change the tab color of the worksheet named "Data" to teal and then saves the workbook as Output.xlsx. | Show how to apply a custom Color (e.g., teal) to an Excel sheet tab with Aspose.Cells in a .NET console application.
// Common Searches: asp.net set worksheet tab color teal using Aspose.Cells | c# Aspose.Cells change Excel sheet tab color programmatically | how to apply custom tab color to an Excel worksheet with Aspose.Cells .NET | example of setting worksheet tab color in Aspose.Cells C#
// Tags: Aspose.Cells set worksheet tab color | C# change Excel sheet tab color | apply teal tab color to worksheet | Aspose.Cells workbook tab customization | Excel worksheet tab color .NET

using Aspose.Cells;
using System.Drawing;

// The example creates a new workbook, renames the first worksheet to "Data", sets its tab color to teal using Aspose.Cells, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet (or add a new one if needed)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Set the worksheet tab color to teal
        sheet.TabColor = Color.Teal;

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
