// Title: Animate Exported Worksheet SVG with SMIL Stroke‑Dashoffset Using Aspose.Cells for .NET
// Description: C# example that creates a workbook, draws a freeform triangle, exports the first worksheet to SVG with Aspose.Cells, loads the SVG via XmlDocument, ensures each <path> has stroke/fill attributes, injects an SMIL <animate> element to animate stroke-dashoffset over 5 seconds (repeat indefinitely), and saves the animated SVG file.
// Keywords: Aspose.Cells SVG animation | SMIL animate C# | stroke-dashoffset Aspose.Cells | export Excel to animated SVG .NET | modify SVG path elements C# | freeform shape SVG Aspose | dynamic SVG from Excel | XmlDocument SVG manipulation | web‑ready animated SVG | C# .NET Excel to SVG
// Common Searches: how to add SMIL animation to SVG exported by Aspose.Cells | C# animate worksheet SVG paths | Aspose.Cells export Excel to animated SVG | add <animate> element to SVG paths with C# | stroke-dashoffset animation for Excel shapes
// Developer Intent: Inject SMIL <animate> elements into SVG paths produced by Aspose.Cells export.
// Use Cases: Create animated diagrams from Excel data for web pages or presentations. | Enhance chart or shape exports with stroke‑dashoffset effects without manual SVG editing. | Automate post‑processing of worksheet SVGs to add dynamic visual cues in dashboards.
// AI Prompts: Generate C# code that uses Aspose.Cells to export a worksheet to SVG and adds an SMIL <animate> element to each <path> for stroke‑dashoffset animation. | Show how to ensure every exported SVG path has stroke and fill attributes before inserting SMIL animation using XmlDocument. | Explain the SvgImageOptions settings (CssPrefix, FitToViewPort) needed for creating an animated SVG from a workbook.

using System;
using System.IO;
using System.Xml;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgAnimation
{
    // C# example that creates a workbook, draws a freeform triangle, exports the first worksheet to SVG with Aspose.Cells, loads the SVG via XmlDocument, ensures each <path> has stroke/fill attributes, injects an SMIL <animate> element to animate stroke-dashoffset over 5 seconds (repeat indefinitely), and saves the animated SVG file.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            ShapeCollection shapes = worksheet.Shapes;

            // 2. Define a simple freeform shape (a triangle) using ShapePath
            ShapePath trianglePath = new ShapePath();
            trianglePath.MoveTo(50, 10);          // top vertex
            trianglePath.LineTo(90, 90);          // bottom right
            trianglePath.LineTo(10, 90);          // bottom left
            trianglePath.Close();                 // close the triangle

            // 3. Add the freeform shape to the worksheet
            // Parameters: topRow, top, leftColumn, left, height, width, paths[]
            shapes.AddFreeform(5, 0, 5, 0, 200, 200, new ShapePath[] { trianglePath });

            // 4. Export the worksheet to SVG using SheetRender and SvgImageOptions
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                ImageType = ImageType.Svg,
                FitToViewPort = true,
                CssPrefix = "anim-"
            };
            SheetRender renderer = new SheetRender(worksheet, svgOptions);
            string svgPath = "worksheet.svg";
            renderer.ToImage(0, svgPath);

            // 5. Load the generated SVG file for manipulation
            XmlDocument svgDoc = new XmlDocument();
            svgDoc.Load(svgPath);

            // 6. Register the SVG namespace for proper node selection
            XmlNamespaceManager nsMgr = new XmlNamespaceManager(svgDoc.NameTable);
            nsMgr.AddNamespace("svg", "http://www.w3.org/2000/svg");

            // 7. Find all <path> elements and add SMIL animation to each
            XmlNodeList pathNodes = svgDoc.SelectNodes("//svg:path", nsMgr);
            foreach (XmlElement pathElem in pathNodes)
            {
                // Ensure the path has a stroke so the animation is visible
                if (string.IsNullOrEmpty(pathElem.GetAttribute("stroke")))
                {
                    pathElem.SetAttribute("stroke", "black");
                }
                if (string.IsNullOrEmpty(pathElem.GetAttribute("fill")))
                {
                    pathElem.SetAttribute("fill", "none");
                }

                // Create an <animate> element that animates stroke-dashoffset
                XmlElement animateElem = svgDoc.CreateElement("animate", "http://www.w3.org/2000/svg");
                animateElem.SetAttribute("attributeName", "stroke-dashoffset");
                animateElem.SetAttribute("from", "100");
                animateElem.SetAttribute("to", "0");
                animateElem.SetAttribute("dur", "5s");
                animateElem.SetAttribute("repeatCount", "indefinite");

                // Append the animation to the path element
                pathElem.AppendChild(animateElem);
            }

            // 8. Save the modified SVG back to disk
            string animatedSvgPath = "worksheet_animated.svg";
            using (FileStream fs = new FileStream(animatedSvgPath, FileMode.Create, FileAccess.Write))
            {
                svgDoc.Save(fs);
            }

            // 9. Optionally, save the workbook (demonstrating the lifecycle rule)
            workbook.Save("AnimatedWorkbook.xlsx");

            Console.WriteLine("SVG animation added successfully.");
        }
    }
}
