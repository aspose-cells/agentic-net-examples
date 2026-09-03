// Title: How to set a worksheet's tab color to blue using Aspose.Cells for .NET (C#)
// AI Prompts: Set the TabColor property of a specific worksheet to System.Drawing.Color.Blue with Aspose.Cells in C#. | Apply a blue tab color to multiple worksheets based on their index using the Aspose.Cells API. | Implement conditional tab coloring (e.g., blue for the first sheet) for Excel workbooks with Aspose.Cells in .NET.
// Common Searches: aspocells c# set worksheet tab color to blue example | how to change Excel sheet tab color programmatically with Aspose.Cells | C# Aspose.Cells change tab color of first worksheet | set tab color for Excel worksheet using Aspose.Cells .NET | Aspose.Cells TabColor property usage tutorial
// Tags: set worksheet tab color Aspose.Cells | Aspose.Cells TabColor property C# | blue worksheet tab Excel .NET | programmatic Excel tab coloring Aspose | change Excel sheet tab color using Aspose.Cells

using Aspose.Cells;
using System.Drawing;

// Creates a workbook, accesses the first worksheet, sets its TabColor to blue, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one using the provided load rule)
        Workbook workbook = new Workbook(); // create rule applied here

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the worksheet tab color to blue
        sheet.TabColor = Color.Blue; // blue tab for visual distinction

        // Save the workbook (using the provided save rule)
        workbook.Save("output.xlsx"); // save rule applied here
    }
}
