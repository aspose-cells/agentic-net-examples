// Title: Read ODS Page Background Color with Aspose.Cells for .NET
// Description: Shows how to load an ODS workbook, access a worksheet's OdsPageBackground via PageSetup, determine if the background is a solid color, and output the Color value or its type.
// Keywords: Aspose.Cells | ODS | page background | background color | C# | .NET | OdsPageBackground | retrieve | extract | audit
// Common Searches: Aspose.Cells read ODS page background color | C# get ODS page background type | How to check ODS page background in .NET | Retrieve ODS page background using Aspose.Cells | Extract solid color from ODS page background
// Developer Intent: Obtain the solid‑color page background defined in an ODS workbook.
// Use Cases: Validate that ODS documents use the corporate brand color for page backgrounds. | Create an inventory of background colors across a batch of ODS files for compliance reporting. | Log the background color before converting ODS to PDF to ensure visual fidelity.
// AI Prompts: Provide C# code that reads the ODS page background color with Aspose.Cells and gracefully handles image or pattern backgrounds. | Explain how to loop through all worksheets in a workbook and list each sheet's ODS page background color. | Show how to set a new solid‑color background for an ODS page and save the changes using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Shows how to load an ODS workbook, access a worksheet's OdsPageBackground via PageSetup, determine if the background is a solid color, and output the Color value or its type.
class Program
{
    static void Main()
    {
        // Load the ODS workbook (replace with your file path)
        Workbook workbook = new Workbook("input.ods");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the ODS page background object
        OdsPageBackground background = worksheet.PageSetup.ODSPageBackground;

        // Determine if the background is set to a color and output it
        if (background.Type == OdsPageBackgroundType.Color)
        {
            Color bgColor = background.Color;
            Console.WriteLine($"ODS page background color: {bgColor}");
        }
        else
        {
            Console.WriteLine($"ODS page background type is not a color. Current type: {background.Type}");
        }
    }
}
