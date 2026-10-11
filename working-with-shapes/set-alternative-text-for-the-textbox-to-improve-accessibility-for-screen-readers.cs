// Title: How to set alternative (alt) text for a textbox shape in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Create an Excel workbook, add a textbox shape to a worksheet, and assign its AlternativeText property for screen‑reader accessibility using Aspose.Cells in C#. | Generate code that inserts a textbox into a worksheet and sets descriptive alt text to improve accessibility compliance with Aspose.Cells for .NET. | Write a C# snippet that saves an Excel file after configuring the AlternativeText of a textbox shape to provide accessibility metadata.
// Common Searches: Aspose.Cells C# set alt text for textbox shape in Excel workbook | How to add accessibility description to Excel shapes using Aspose.Cells .NET | Example of setting AlternativeText property on a textbox in Aspose.Cells
// Tags: Aspose.Cells textbox alt description | C# shape accessibility in Excel | Excel textbox alt text property | Aspose.Cells add accessible shape | C# generate accessible Excel workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Creates a new workbook, inserts a textbox shape on the first worksheet, assigns AlternativeText for screen‑reader accessibility, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape to the worksheet
            // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
            TextBox textbox = sheet.Shapes.AddTextBox(1, 1, 0, 0, 200, 100);

            // Set alternative text to improve accessibility for screen readers
            textbox.AlternativeText = "Summary of quarterly sales figures";

            // Save the workbook
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
