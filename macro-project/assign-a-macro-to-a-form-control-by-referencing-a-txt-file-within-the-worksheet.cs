using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AssignMacroToFormControl
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Enable macros for the workbook (required for macro-enabled files)
        workbook.Settings.EnableMacros = true;

        // Path to the text file that contains the macro name (e.g., "MyMacro()")
        string macroFilePath = "Macro.txt";

        // Read the macro name from the text file, or use a default if the file is missing
        string macroName = File.Exists(macroFilePath)
            ? File.ReadAllText(macroFilePath).Trim()
            : "MyMacro()";

        // Add a button shape (form control) to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, height, width
        Shape button = sheet.Shapes.AddButton(2, 2, 0, 0, 30, 100);

        // Assign the macro to the button using the MacroName property
        button.MacroName = macroName;

        // Optionally set a caption for the button
        button.Text = "Run Macro";

        // Save the workbook as a macro-enabled file
        workbook.Save("WorkbookWithMacroButton.xlsm", SaveFormat.Xlsm);
    }
}