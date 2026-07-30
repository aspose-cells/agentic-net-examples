// Title: C# – Using Aspose.Cells ShowFormulas to Display Cell Formulas for Debugging
// Description: Learn how to toggle the ShowFormulas property in Aspose.Cells for .NET. This example creates a workbook, writes a formula to A1, prints the calculated value, switches ShowFormulas on to display the formula text, and saves the file. Ideal for developers who need to inspect or debug complex Excel calculations programmatically.
// Keywords: Aspose.Cells ShowFormulas | display formulas C# | debug Excel formulas Aspose | toggle formula view .NET | Aspose.Cells worksheet debugging | C# Excel formula display | Aspose.Cells sample code
// Common Searches: how to enable ShowFormulas in Aspose.Cells | display formulas instead of values Aspose.Cells C# | debug Excel calculations with Aspose.Cells | toggle formula view in workbook using Aspose | Aspose.Cells ShowFormulas example GitHub
// Developer Intent: Show formulas in a worksheet instead of calculated results to aid debugging.
// Use Cases: Set worksheet.ShowFormulas = true while testing to verify complex formulas. | Turn ShowFormulas off before publishing a workbook so end users see values. | Save a workbook with ShowFormulas enabled to open in Excel with formulas displayed.
// AI Prompts: Generate C# code that toggles Aspose.Cells ShowFormulas on a specific worksheet and explains the output difference. | Create an example that iterates through all worksheets in a workbook and enables ShowFormulas for debugging. | Explain how to programmatically verify that the ShowFormulas setting persists after saving and reopening the workbook.

using System;
using Aspose.Cells;

// Learn how to toggle the ShowFormulas property in Aspose.Cells for .NET. This example creates a workbook, writes a formula to A1, prints the calculated value, switches ShowFormulas on to display the formula text, and saves the file. Ideal for developers who need to inspect or debug complex Excel calculations programmatically.
class ShowFormulasDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Place a formula in cell A1
        worksheet.Cells["A1"].Formula = "=1+2+3";

        // Show the calculated result (default behavior)
        worksheet.ShowFormulas = false;
        Console.WriteLine("ShowFormulas OFF: " + worksheet.Cells["A1"].StringValue);

        // Enable formula view for debugging
        worksheet.ShowFormulas = true;
        Console.WriteLine("ShowFormulas ON: " + worksheet.Cells["A1"].StringValue);

        // Save the workbook (optional, to verify the setting in the file)
        workbook.Save("ShowFormulasDemo.xlsx");
    }
}
