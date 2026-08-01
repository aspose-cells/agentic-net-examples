// Title: Aspose.Cells for .NET – Add and Modify Shape Adjustment Guides with VBA Support
// Description: Step‑by‑step guide that creates a workbook, inserts a RightArrowCallout auto shape, works with ShapeGuideCollection to add and change guide values, optionally embeds a VBA class module, and saves the file as XLSX and XLSM using Aspose.Cells for .NET.
// Keywords: Aspose.Cells shape guides | C# auto shape adjustment | ShapeGuideCollection example | add VBA module Aspose.Cells | save workbook as XLSM | RightArrowCallout geometry | Aspose.Cells .NET tutorial | modify shape geometry C#
// Common Searches: how to set shape adjustment guides in Aspose.Cells C# | add RightArrowCallout auto shape Aspose.Cells | update shape guide values programmatically | embed VBA macro in Aspose.Cells workbook | save Aspose.Cells workbook with macros
// Developer Intent: Add an auto shape, configure its adjustment guides, optionally attach VBA code, and export the workbook in both XLSX and macro‑enabled XLSM formats.
// Use Cases: Create a callout shape and define custom adj1‑adj4 values to control arrow dimensions. | Programmatically change all guide values to resize the shape uniformly and verify changes via console output. | Insert a VBA class module containing a simple macro and generate an XLSM file that preserves the macro.
// AI Prompts: Generate C# code that adds a RoundedRectangle auto shape, defines four custom adjustment guides, and saves the workbook as XLSX using Aspose.Cells. | Show how to read existing shape guide values from a workbook, modify them based on user input, and persist the changes. | Provide an example of adding a standard VBA module with a macro to a workbook and saving it as an XLSM file with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Vba;

namespace AsposeCellsGuide
{
    // Step‑by‑step guide that creates a workbook, inserts a RightArrowCallout auto shape, works with ShapeGuideCollection to add and change guide values, optionally embeds a VBA class module, and saves the file as XLSX and XLSM using Aspose.Cells for .NET.
    public class ShapeGuideDeveloperGuide
    {
        public static void Main()
        {
            // Step 1: Create a new workbook and obtain the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Step 2: Add an auto shape that supports adjustment guides (e.g., RightArrowCallout).
            Shape shape = worksheet.Shapes.AddAutoShape(AutoShapeType.RightArrowCallout, 2, 0, 2, 0, 200, 150);

            // Step 3: Retrieve the ShapeGuideCollection from the shape's geometry adjustments.
            ShapeGuideCollection guides = shape.Geometry.ShapeAdjustValues;

            // Step 4: Add adjustment guides using the Add(string name, double value) method.
            guides.Add("adj1", 25.5);
            guides.Add("adj2", 30.0);
            guides.Add("adj3", 25.5);
            guides.Add("adj4", 35.0);

            // Step 5: Read and display the initial guide values via the indexer.
            Console.WriteLine("Initial guide values:");
            for (int i = 0; i < guides.Count; i++)
            {
                Console.WriteLine($"Guide {i + 1}: {guides[i].Value}");
            }

            // Step 6: Modify the guide values.
            guides[0].Value = 20.0;
            guides[1].Value = 20.0;
            guides[2].Value = 20.0;
            guides[3].Value = 20.0;

            // Step 7: Verify and display the updated guide values.
            Console.WriteLine("Updated guide values:");
            for (int i = 0; i < guides.Count; i++)
            {
                Console.WriteLine($"Guide {i + 1}: {guides[i].Value}");
            }

            // Optional Step: Add a VBA module to demonstrate VbaModuleType usage.
            VbaProject vbaProject = workbook.VbaProject;
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "DemoModule");
            VbaModule vbaModule = vbaProject.Modules[moduleIndex];
            vbaModule.Codes = "Sub ShowMessage()\r\n    MsgBox \"Aspose.Cells guide executed\"\r\nEnd Sub";

            // Step 8: Save the workbook.
            // Save as XLSX (shape only) and XLSM (including VBA).
            workbook.Save("ShapeGuideDeveloperGuide.xlsx");
            workbook.Save("ShapeGuideDeveloperGuide_WithVba.xlsm", SaveFormat.Xlsm);
        }
    }
}
