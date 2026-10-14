// Title: How to apply a blue tab color to a worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that applies a blue tab color to a specific worksheet with Aspose.Cells and saves the file. | Show an example of using the Aspose.Cells API to change a worksheet's tab color to blue in a .NET project. | Provide step‑by‑step instructions for setting a custom tab color for an Excel sheet via Aspose.Cells in C#.
// Common Searches: asp.net change worksheet tab color to blue using aspose.cells | c# example for setting excel sheet tab color with Aspose.Cells library | how to customize worksheet tab colors programmatically in Aspose.Cells for .NET | Aspose.Cells set custom tab color for first worksheet C# code sample
// Tags: Aspose.Cells apply tab color C# | Excel worksheet custom tab color .NET | blue tab color Aspose.Cells example | worksheet tab styling using Aspose.Cells | C# set worksheet tab hue Aspose

using Aspose.Cells;
using System.Drawing;

// // Creates a new workbook, accesses the first worksheet, assigns its TabColor property to Color.Blue, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the worksheet's tab color to blue
        sheet.TabColor = Color.Blue;

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
