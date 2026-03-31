using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

public class BindMacroToControl
{
    public static void Main()
    {
        Run();
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape that will act as a button
        // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, width, height
        Shape button = sheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 30);

        // Assign a macro to the shape
        button.MacroName = "MyMacro";

        // Link the button's value to cell A1
        button.LinkedCell = "$A$1";

        // Set an initial numeric value in the linked cell
        Cell linkedCell = sheet.Cells["A1"];
        linkedCell.PutValue(1234.56);

        // Create a style with a built‑in number format (2 corresponds to "0.00")
        Style numberStyle = workbook.CreateStyle();
        numberStyle.Number = 2; // Built‑in format "0.00"

        // Configure a StyleFlag to apply only the number format
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;

        // Apply the number format style to the linked cell
        AsposeRange range = sheet.Cells.CreateRange("A1");
        range.ApplyStyle(numberStyle, flag);

        // Save the workbook
        workbook.Save("MacroControlWithNumberFormat.xlsx");
    }
}