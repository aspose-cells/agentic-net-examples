// Title: Add SMIL Stroke‑Dashoffset Animation to Aspose.Cells‑Generated Worksheet SVG (C#)
// Description: Creates a workbook, renders the first worksheet to SVG with Aspose.Cells, loads the SVG, injects an <animate> element into every <path> to animate the stroke‑dashoffset over 5 seconds (repeat indefinitely), saves the animated file, and removes the temporary SVG.
// Keywords: Aspose.Cells SVG animation | C# SMIL animate path | stroke-dashoffset SVG | SheetRender to SVG | modify Aspose.Cells SVG | dynamic Excel chart SVG | .NET SVG manipulation | inject <animate> into SVG
// Common Searches: how to add SMIL animation to SVG exported by Aspose.Cells | C# code to animate path elements in worksheet SVG | inject <animate> tag into Aspose.Cells generated SVG | animated Excel chart SVG using Aspose.Cells | Aspose.Cells render worksheet as animated SVG
// Developer Intent: Insert SMIL <animate> elements into each <path> of an SVG produced by Aspose.Cells to create continuous stroke‑dashoffset animation.
// Use Cases: Display a live‑drawing sales chart on a web dashboard. | Generate SVG reports where trend lines are highlighted with moving dashes. | Create presentation‑ready Excel‑to‑SVG conversions that include built‑in animations.
// AI Prompts: Generate C# code that loads an Aspose.Cells SVG file and adds an <animate> element to all <path> nodes to animate stroke-dashoffset for 5 seconds with infinite repeat. | Explain the process of rendering a worksheet to SVG with Aspose.Cells, then enhancing the SVG with SMIL animation while preserving viewport settings. | Provide step‑by‑step instructions to create a temporary SVG from a workbook, inject SMIL <animate> tags, save the animated SVG, and clean up temporary files.

using System;
using System.IO;
using System.Xml;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgAnimation
{
    // Creates a workbook, renders the first worksheet to SVG with Aspose.Cells, loads the SVG, injects an <animate> element into every <path> to animate the stroke‑dashoffset over 5 seconds (repeat indefinitely), saves the animated file, and removes the temporary SVG.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and populate sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(210);
                sheet.Cells["B4"].PutValue(150);

                // 2. Render the worksheet to an intermediate SVG file
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    // SVG output is implicit; set viewport fitting
                    FitToViewPort = true
                };

                string tempSvgPath = "worksheet_temp.svg";

                // Render the first worksheet page to a temporary SVG file
                SheetRender renderer = new SheetRender(sheet, svgOptions);
                renderer.ToImage(0, tempSvgPath);

                // 3. Load the generated SVG for manipulation
                if (!File.Exists(tempSvgPath))
                    throw new FileNotFoundException("Temporary SVG file was not created.", tempSvgPath);

                XmlDocument svgDoc = new XmlDocument();
                svgDoc.Load(tempSvgPath);

                // Define SVG namespace for node selection
                XmlNamespaceManager nsMgr = new XmlNamespaceManager(svgDoc.NameTable);
                nsMgr.AddNamespace("svg", "http://www.w3.org/2000/svg");

                // 4. Inject SMIL <animate> elements into each <path>
                XmlNodeList pathNodes = svgDoc.SelectNodes("//svg:path", nsMgr);
                if (pathNodes != null)
                {
                    foreach (XmlNode pathNode in pathNodes)
                    {
                        XmlElement animateElem = svgDoc.CreateElement("animate", "http://www.w3.org/2000/svg");
                        animateElem.SetAttribute("attributeName", "stroke-dashoffset");
                        animateElem.SetAttribute("from", "0");
                        animateElem.SetAttribute("to", "100");
                        animateElem.SetAttribute("dur", "5s");
                        animateElem.SetAttribute("repeatCount", "indefinite");
                        pathNode.AppendChild(animateElem);
                    }
                }

                // 5. Save the modified SVG with animation
                string animatedSvgPath = "worksheet_animated.svg";
                using (FileStream fs = new FileStream(animatedSvgPath, FileMode.Create, FileAccess.Write))
                {
                    svgDoc.Save(fs);
                }

                // Clean up temporary file
                if (File.Exists(tempSvgPath))
                    File.Delete(tempSvgPath);

                Console.WriteLine($"Animated SVG saved to: {Path.GetFullPath(animatedSvgPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
