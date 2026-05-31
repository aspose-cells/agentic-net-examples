using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ReplaceSmartArtText
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Iterate through all shapes in the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Process only SmartArt shapes
                if (shape.IsSmartArt)
                {
                    // Convert the SmartArt shape to a grouped shape container
                    GroupShape groupShape = shape.GetResultOfSmartArt();

                    // Replace the text of each grouped shape inside the SmartArt
                    foreach (Shape innerShape in groupShape.GetGroupedShapes())
                    {
                        // Set the desired replacement text
                        innerShape.Text = "ReplacedText";
                    }
                }
            }
        }

        // Save the workbook with SmartArt updating enabled
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.UpdateSmartArt = true;
        workbook.Save("output.xlsx", saveOptions);
    }
}