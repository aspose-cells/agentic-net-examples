using System;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMacroAssignment
{
    public class MacroInfo
    {
        public string MacroName { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            string json = @"{ ""MacroName"": ""MyMacro"" }";
            MacroInfo macroInfo = JsonSerializer.Deserialize<MacroInfo>(json)!;

            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a Forms button directly
            Button button = sheet.Shapes.AddButton(
                2, // upperLeftRow
                2, // upperLeftColumn
                0, // upperLeftRowOffset
                0, // upperLeftColumnOffset
                100, // width
                30   // height
            );

            // Note: The Button class in the current Aspose.Cells version does not expose a Macro property.
            // If macro assignment is required, ensure you are using a version that supports it.
            // button.Macro = macroInfo.MacroName; // Uncomment when supported

            button.Text = "Run Macro";

            workbook.Save("WorkbookWithMacroButton.xlsx");
        }
    }
}