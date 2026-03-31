using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Text.Json;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a button shape (Forms button)
        Button button = sheet.Shapes.AddButton(2, 2, 0, 0, 100, 30);
        button.Name = "MyButton";
        button.Text = "Run Macro";

        // Define the macro name (cannot be set directly if the API version lacks the property)
        string macroName = "MyMacro";

        // Build a JSON mapping that links the control name to its macro
        var mapping = new
        {
            ControlName = button.Name,
            Macro = macroName
        };
        string jsonMapping = JsonSerializer.Serialize(mapping, new JsonSerializerOptions { WriteIndented = true });

        // Store the JSON mapping in a worksheet cell (optional, just for illustration)
        sheet.Cells["A1"].PutValue(jsonMapping);

        // Save the workbook
        workbook.Save("MacroButton.xlsx");
    }
}