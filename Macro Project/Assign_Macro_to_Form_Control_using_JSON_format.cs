using System;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a Forms button to the worksheet (row 1, column 1, no offset, width 100, height 30)
        // The AddButton method creates a button shape that can be linked to a macro.
        Shape button = worksheet.Shapes.AddButton(1, 1, 0, 0, 100, 30);

        // JSON string that defines the macro name to be assigned to the button
        string json = "{\"MacroName\":\"DoWork()\"}";

        // Parse the JSON and extract the macro name
        using (JsonDocument doc = JsonDocument.Parse(json))
        {
            JsonElement root = doc.RootElement;
            if (root.TryGetProperty("MacroName", out JsonElement macroElement))
            {
                string macroName = macroElement.GetString();

                // Assign the macro name to the button shape
                button.MacroName = macroName;
                Console.WriteLine($"Macro assigned to button: {button.MacroName}");
            }
            else
            {
                Console.WriteLine("MacroName property not found in JSON.");
            }
        }

        // Save the workbook to a file
        workbook.Save("ButtonWithMacro.xlsx");
    }
}