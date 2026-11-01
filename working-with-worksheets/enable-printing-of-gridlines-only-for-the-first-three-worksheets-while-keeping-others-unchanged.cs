// Title: How to enable printing of gridlines only on the first three worksheets in an Excel file using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, sets the PageSetup.PrintGridlines flag to true for the first three worksheets, and saves the updated file. | Generate a C# loop using Aspose.Cells that applies the gridline‑printing setting to selected sheets while leaving the rest unchanged. | Create a .NET snippet that modifies an existing Excel file so that only the initial three sheets print gridlines when printed.
// Common Searches: Aspose.Cells C# enable gridline printing for specific sheets in a workbook | how to apply page setup options to the first three worksheets using Aspose.Cells | C# loop through Excel worksheets and turn on gridlines for selected sheets with Aspose.Cells | set print options per worksheet Aspose.Cells .NET example
// Tags: Aspose.Cells set PrintGridlines per worksheet | C# enable gridline printing for selected Excel sheets | Aspose.Cells page setup gridlines first three sheets | modify workbook print options Aspose.Cells .NET | loop worksheets Aspose.Cells PrintGridlines

using Aspose.Cells;
using System;

// Loads an existing Excel workbook, enables gridline printing on the first three worksheets by setting PageSetup.PrintGridlines to true, and saves the modified file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        string inputPath = @"C:\Input\Sample.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Enable printing of gridlines for the first three worksheets
        for (int i = 0; i < workbook.Worksheets.Count && i < 3; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];
            // Set the PrintOptions to print gridlines
            sheet.PageSetup.PrintGridlines = true;
        }

        // Save the modified workbook (replace with your desired output path)
        string outputPath = @"C:\Output\Sample_GridlinesEnabled.xlsx";
        workbook.Save(outputPath);
    }
}
