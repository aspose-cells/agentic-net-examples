// Title: How to save a macro‑enabled Excel workbook after assigning a VBA macro to a button using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells add a button to a worksheet and keep the workbook macro‑enabled when saving | Saving an .xlsm file after setting button.MacroName with Aspose.Cells .NET | Assign VBA macro to shape in Aspose.Cells and preserve macros on save | Use OoxmlSaveOptions to save macro‑enabled workbook after adding button in C# | Aspose.Cells preserve existing VBA macros while adding new button shape
// Tags: button shape macro assignment Aspose.Cells | OoxmlSaveOptions for .xlsm output | VBA macro binding to worksheet shape C# | retain VBA macros during Aspose.Cells save | macro‑enabled Excel file generation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Saving;

// The example loads an existing .xlsm workbook, adds a button shape to the first worksheet, sets its caption, assigns the VBA macro "MyMacro" to the button, and then saves the workbook as a macro‑enabled file using OoxmlSaveOptions with SaveFormat.Xlsm, while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsm";
            const string outputPath = "output.xlsm";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load an existing macro‑enabled workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a button shape to the worksheet.
            // Parameters: upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, height, width
            Button button = sheet.Shapes.AddButton(2, 1, 2, 1, 30, 100);

            // Set the button caption
            button.Text = "Run Macro";

            // Assign a macro to the button (the macro must exist in the VBA project)
            button.MacroName = "MyMacro";

            // Save the workbook as a macro‑enabled file
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsm);
            workbook.Save(outputPath, saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
