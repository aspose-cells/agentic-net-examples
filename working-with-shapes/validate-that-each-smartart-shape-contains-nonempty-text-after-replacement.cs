// Title: C# – Validate and Replace Empty Text in SmartArt Shapes with Aspose.Cells
// Description: Load an Excel workbook, walk through every worksheet, locate SmartArt shapes, convert them to GroupShape via GetResultOfSmartArt, insert a default string where text is blank, confirm the replacement succeeded, and save the file with UpdateSmartArt enabled using OoxmlSaveOptions.
// Keywords: Aspose.Cells SmartArt C# | replace empty SmartArt text | GetResultOfSmartArt example | UpdateSmartArt save option | .NET Excel SmartArt validation | grouped shapes Aspose
// Common Searches: how to fill missing SmartArt labels with Aspose.Cells | C# code to check SmartArt nodes for empty text | Aspose.Cells replace blank SmartArt text in Excel | save workbook with updated SmartArt using OoxmlSaveOptions
// Developer Intent: Automatically ensure that every SmartArt node in an Excel file contains non‑empty text after applying a default value.
// Use Cases: Clean up template files before distribution so all SmartArt elements are labeled. | Run a quality‑control step in a reporting pipeline that verifies SmartArt completeness. | Prepare Excel workbooks for downstream data extraction where empty SmartArt nodes cause errors.
// AI Prompts: Write C# code with Aspose.Cells that iterates through SmartArt shapes and sets "DefaultText" for any empty node. | Show how to throw a detailed exception when a SmartArt shape remains empty after replacement. | Explain the role of OoxmlSaveOptions.UpdateSmartArt when persisting changes to a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtValidation
{
    // Load an Excel workbook, walk through every worksheet, locate SmartArt shapes, convert them to GroupShape via GetResultOfSmartArt, insert a default string where text is blank, confirm the replacement succeeded, and save the file with UpdateSmartArt enabled using OoxmlSaveOptions.
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets and their shapes
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        // Convert the SmartArt to a GroupShape to access its constituent shapes
                        GroupShape group = shape.GetResultOfSmartArt();

                        if (group != null)
                        {
                            // Iterate through each shape inside the SmartArt
                            foreach (Shape smartShape in group.GetGroupedShapes())
                            {
                                // If the text is empty or whitespace, replace it with a default value
                                if (string.IsNullOrWhiteSpace(smartShape.Text))
                                {
                                    smartShape.Text = "DefaultText";
                                }

                                // Validate that the text is now non‑empty
                                if (string.IsNullOrWhiteSpace(smartShape.Text))
                                {
                                    throw new InvalidOperationException(
                                        $"SmartArt shape '{smartShape.Name}' contains empty text after replacement.");
                                }
                            }
                        }
                    }
                }
            }

            // Save the workbook with SmartArt updating enabled
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.UpdateSmartArt = true; // ensures SmartArt changes are persisted
            workbook.Save("output.xlsx", saveOptions);
        }
    }
}
