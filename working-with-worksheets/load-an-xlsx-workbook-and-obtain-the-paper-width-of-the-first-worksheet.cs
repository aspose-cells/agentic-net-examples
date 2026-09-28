// Title: Get the paper width of the first worksheet in an XLSX workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Provide C# code that opens an XLSX file with Aspose.Cells and returns the PageSetup.PaperWidth of the first sheet. | Show how to read the print page width of a worksheet using Aspose.Cells in a .NET console application. | Modify the example to also display the paper height alongside the width for the first worksheet.
// Common Searches: Aspose.Cells C# retrieve first sheet paper width from XLSX | How to read worksheet page setup dimensions using Aspose.Cells .NET | Get print page width of a worksheet in C# with Aspose.Cells | C# Aspose.Cells example for accessing PageSetup.PaperWidth property | Read paper size settings of the first worksheet in an Excel file using Aspose.Cells
// Tags: Aspose.Cells read worksheet paper width | C# page setup dimensions XLSX | retrieve print width property Aspose | first worksheet page setup Aspose.Cells | access PaperWidth API .NET

using System;
using Aspose.Cells;

// Loads an XLSX workbook, selects the first worksheet, reads its PageSetup.PaperWidth (points), and writes the value to the console.
class Program
{
    static void Main()
    {
        // Load the XLSX workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (index 0)
        Worksheet firstSheet = workbook.Worksheets[0];

        // Obtain the paper width of the worksheet (in points)
        double paperWidth = firstSheet.PageSetup.PaperWidth;

        // Output the paper width
        Console.WriteLine($"Paper width of the first worksheet: {paperWidth} points");
    }
}
