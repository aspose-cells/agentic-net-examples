// Title: Hide the second worksheet of an existing Excel workbook and save the changes using Aspose.Cells for .NET (C#)
// AI Prompts: Load an .xlsx file with Aspose.Cells, set the worksheet at index 1 to hidden by changing its IsVisible property, and save the workbook to a new file. | Write C# code that opens a workbook, hides the second sheet using Aspose.Cells, and writes the result as output.xlsx.
// Common Searches: Aspose.Cells C# hide worksheet at index 1 and save workbook | How to set a specific sheet to hidden in an existing Excel file using Aspose.Cells .NET | C# example for hiding the second sheet in an Excel workbook with Aspose.Cells | Change worksheet visibility to hidden without opening Excel using Aspose.Cells | Save workbook after modifying sheet visibility with Aspose.Cells for .NET
// Tags: worksheet IsVisible false Aspose.Cells | hide specific sheet .NET Aspose.Cells | modify workbook sheet visibility C# | save workbook after sheet hide Aspose.Cells | second worksheet hidden Aspose.Cells example

using Aspose.Cells;
using System;

// // Loads 'input.xlsx', hides the second worksheet by setting IsVisible = false, and saves the updated workbook as 'output.xlsx' using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the second worksheet (index 1) and hide it
        Worksheet secondSheet = workbook.Worksheets[1];
        secondSheet.IsVisible = false; // Alternatively: secondSheet.Visibility = VisibilityType.Hidden;

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
