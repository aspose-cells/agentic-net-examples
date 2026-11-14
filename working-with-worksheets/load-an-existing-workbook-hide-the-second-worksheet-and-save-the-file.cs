// Title: Hide the second worksheet in an existing Excel workbook and save the file with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells hide specific worksheet and save workbook | How to make the second sheet invisible in an Excel file using Aspose.Cells | Aspose.Cells set worksheet visibility to hidden programmatically in .NET | Save a workbook after changing worksheet visibility with Aspose.Cells C# | Hide a worksheet by index using Aspose.Cells for .NET
// Tags: hide worksheet Aspose.Cells | worksheet visibility C# | modify worksheet visibility Aspose.Cells | save workbook after changes Aspose.Cells | Excel sheet hidden by index .NET

using Aspose.Cells;
using System;

// Loads 'input.xlsx', hides the second worksheet (index 1) by setting its visibility to hidden, and saves the updated workbook as 'output.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Hide the second worksheet (index 1, zero‑based)
        Worksheet secondSheet = workbook.Worksheets[1];
        secondSheet.IsVisible = false;   // Alternatively: secondSheet.Visibility = VisibilityType.Hidden;

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
