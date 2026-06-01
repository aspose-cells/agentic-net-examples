using System;
using Aspose.Cells;
using System.Drawing;

class InheritRowBackground
{
    static void Main()
    {
        // Load the template workbook that contains the desired row shading scheme
        Workbook templateWb = new Workbook("Template.xlsx");
        Worksheet templateWs = templateWb.Worksheets[0];

        // Index of the template row that has the background color you want to inherit (zero‑based)
        int templateRowIdx = 1; // e.g., second row in the template

        // ------------------------------------------------------------
        // Option 1 – Use Row.CopySettings (copies style, height, visibility, etc.)
        // ------------------------------------------------------------
        Row sourceRow = templateWs.Cells.Rows[templateRowIdx];

        // Create a new workbook where rows will be generated
        Workbook resultWb = new Workbook();
        Worksheet resultWs = resultWb.Worksheets[0];
        Cells resultCells = resultWs.Cells;

        // Generate 5 rows and inherit the background color from the template row
        for (int i = 0; i < 5; i++)
        {
            // Fill some sample data in column A
            resultCells[i, 0].PutValue($"Item {i + 1}");

            // Get the target row object in the result sheet
            Row targetRow = resultWs.Cells.Rows[i];

            // Copy all settings (including background) from the template row
            targetRow.CopySettings(sourceRow, true);
        }

        // ------------------------------------------------------------
        // Option 2 – Use Range.CopyStyle (copies only style information)
        // ------------------------------------------------------------
        // If you prefer to copy only the style (no height/visibility), uncomment the block below
        /*
        for (int i = 0; i < 5; i++)
        {
            // Fill data as before
            resultCells[i, 0].PutValue($"Item {i + 1}");

            // Define source range that represents the whole template row
            Range srcRange = templateWs.Cells.CreateRange(templateRowIdx, 0, 1, templateWs.Cells.MaxColumn + 1);

            // Define destination range that represents the target row in the result sheet
            Range destRange = resultWs.Cells.CreateRange(i, 0, 1, resultWs.Cells.MaxColumn + 1);

            // Copy style (including background color) from source to destination
            destRange.CopyStyle(srcRange);
        }
        */

        // Save the generated workbook
        resultWb.Save("GeneratedWithInheritedShading.xlsx");
    }
}