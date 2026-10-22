// Title: Set the worksheet tab color to teal in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new workbook with Aspose.Cells, sets the first worksheet's tab color to teal, and saves it as an .xlsx file. | Show how to change the tab color of an existing worksheet to teal using the Aspose.Cells API in a .NET application.
// Common Searches: Aspose.Cells C# how to change Excel sheet tab color to teal | programmatically set worksheet tab color using Aspose.Cells .NET | C# code example for setting tab color of a worksheet in Aspose.Cells | change Excel worksheet tab color with Aspose.Cells library | set custom tab color for a worksheet in Aspose.Cells for .NET
// Tags: Aspose.Cells set worksheet tab color | C# worksheet tab color teal | Aspose.Cells workbook tab coloring | Excel .xlsx tab color customization Aspose | Aspose.Cells API TabColor property usage

using Aspose.Cells;
using System.Drawing;

// This example creates a new workbook, accesses the first worksheet, assigns the teal color to its tab via the TabColor property, and saves the workbook as TabColorDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook(); // {CreateWorkbook}

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the worksheet tab color to teal
        sheet.TabColor = Color.Teal;

        // Save the workbook
        workbook.Save("TabColorDemo.xlsx"); // {SaveWorkbook}
    }
}
