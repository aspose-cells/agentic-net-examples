// Title: Set a Solid Blue Page Background in an ODS File with Aspise.Cells for .NET
// Description: Demonstrates how to create a workbook, access the first worksheet's PageSetup, configure OdsPageBackground to a solid blue color, and save the result as an ODS document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells ODS background | C# OdsPageBackground | solid color ODS page | save ODS with background | Aspose.Cells .NET example
// Common Searches: how to set page background color in ODS using Aspose.Cells | C# code for solid blue ODS background | Aspose.Cells OdsPageBackground Type Color example | save ODS file with custom background .NET
// Developer Intent: Apply a solid blue page background to an ODS workbook and persist the file.
// Use Cases: Brand‑consistent ODS reports with a corporate blue background. | Automated generation of printable ODS sheets that need a highlighted background color. | Creating reusable ODS templates pre‑styled with a solid color for downstream data exports.
// AI Prompts: Generate C# code to apply a gradient background to an ODS page with Aspose.Cells. | Show how to change the ODS page background color dynamically based on user input. | Explain the steps to add a background image to an ODS workbook using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsOdsBackgroundDemo
{
    // Demonstrates how to create a workbook, access the first worksheet's PageSetup, configure OdsPageBackground to a solid blue color, and save the result as an ODS document using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet's page setup
            PageSetup pageSetup = workbook.Worksheets[0].PageSetup;

            // Get the ODS page background object
            OdsPageBackground background = pageSetup.ODSPageBackground;

            // Set the background type to solid color
            background.Type = OdsPageBackgroundType.Color;

            // Apply solid blue color
            background.Color = Color.Blue;

            // Save the workbook as ODS with the background applied
            workbook.Save("BlueBackground.ods");
        }
    }
}
