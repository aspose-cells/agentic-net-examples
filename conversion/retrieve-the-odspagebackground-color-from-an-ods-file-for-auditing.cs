// Title: How to read the ODS page background color of a worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .ods file with Aspose.Cells, accesses the first worksheet's PageSetup, checks if OdsPageBackground.Type equals Color, and prints the ARGB components of the background color. | Show how to programmatically determine whether an ODS file's page background is a solid color and retrieve its Color value using the Aspose.Cells OdsPageBackground class in .NET.
// Common Searches: Aspose.Cells C# read ODS page background color | Get ARGB values of ODS page background with Aspose.Cells .NET | Check if ODS page background is a solid color using Aspose.Cells | How to access OdsPageBackground.Type in C# | Retrieve ODS page setup background color programmatically
// Tags: Aspose.Cells ODS page background color extraction | C# read ODS page setup background | ODS file page background type detection | .NET Aspose.Cells ODS background API usage | retrieve ARGB values from ODS background

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Ods;

// The example loads an ODS workbook, accesses the first worksheet's PageSetup, obtains the OdsPageBackground object, verifies that its Type is Color, and then outputs the ARGB components of the background color; if the type is not Color, it reports the actual background type.
class Program
{
    static void Main()
    {
        // Load the existing ODS workbook
        Workbook workbook = new Workbook("input.ods");

        // Access the first worksheet's page setup
        PageSetup pageSetup = workbook.Worksheets[0].PageSetup;

        // Retrieve the ODS page background object
        OdsPageBackground background = pageSetup.ODSPageBackground;

        // Determine if the background is set to a color and output its value
        if (background.Type == OdsPageBackgroundType.Color)
        {
            Color bgColor = background.Color;
            Console.WriteLine($"ODS Page Background Type: Color");
            Console.WriteLine($"Background Color: A={bgColor.A}, R={bgColor.R}, G={bgColor.G}, B={bgColor.B}");
        }
        else
        {
            Console.WriteLine($"ODS Page Background Type is not Color. Current Type: {background.Type}");
        }
    }
}
