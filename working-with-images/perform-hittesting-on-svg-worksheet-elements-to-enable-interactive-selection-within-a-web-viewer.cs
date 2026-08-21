// Title: Aspose.Cells .NET – Render Worksheet to SVG with DrawObject metadata for client‑side hit testing
// Description: This example loads an Excel workbook, configures SvgImageOptions (FitToViewPort), attaches a custom DrawObjectEventHandler to capture each draw object's type, sheet index, page number, and bounding rectangle, saves the worksheet as SVG, and writes the collected data to a JSON file for interactive hit‑testing in a web viewer.
// Keywords: Aspose.Cells | C# SVG rendering | draw object event handler | hit testing Excel SVG | FitToViewPort | worksheet to SVG | JSON map of SVG elements | interactive Excel viewer | client‑side hit test | Aspose.Cells .NET example
// Common Searches: Aspose.Cells export worksheet to SVG with element coordinates | How to capture draw object bounds during SVG rendering in C# | Generate JSON map for SVG hit testing using Aspose.Cells | Enable interactive selection of Excel cells in a web viewer | FitToViewPort option Aspose.Cells SVG
// Developer Intent: Create an SVG of a worksheet and a JSON coordinate map to enable client‑side hit testing of Excel elements.
// Use Cases: Render Excel sheets as SVG and provide a JSON map so JavaScript can highlight cells, charts, or shapes on hover. | Implement clickable charts or tables in a web dashboard that trigger server actions based on SVG element IDs. | Store draw‑object metadata for analytics, such as generating tooltips or exporting selected ranges.
// AI Prompts: Write C# code that reads drawObjectsMap.json and returns the worksheet element at a given mouse coordinate. | Show how to extend SvgDrawObjectHandler to add a unique ID and layer information to the JSON output. | Explain how to integrate worksheet.svg and drawObjectsMap.json into a JavaScript front‑end for hit testing and interactive selection.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgHitTestDemo
{
    // Custom handler to capture draw objects during SVG rendering
    // This example loads an Excel workbook, configures SvgImageOptions (FitToViewPort), attaches a custom DrawObjectEventHandler to capture each draw object's type, sheet index, page number, and bounding rectangle, saves the worksheet as SVG, and writes the collected data to a JSON file for interactive hit‑testing in a web viewer.
    class SvgDrawObjectHandler : DrawObjectEventHandler
    {
        // List to store information about each draw object
        public List<DrawObjectInfo> ObjectsInfo { get; } = new List<DrawObjectInfo>();

        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            // Record relevant data for hit‑testing on the client side
            ObjectsInfo.Add(new DrawObjectInfo
            {
                Type = drawObject.Type.ToString(),
                SheetIndex = drawObject.SheetIndex,
                Page = drawObject.CurrentPage + 1, // make it 1‑based for UI
                X = x,
                Y = y,
                Width = width,
                Height = height
            });
        }
    }

    // Simple DTO for serialization
    class DrawObjectInfo
    {
        public string Type { get; set; }
        public int SheetIndex { get; set; }
        public int Page { get; set; }
        public float X { get; set; }
        public float Y { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Load an existing workbook (replace with your actual file path)
            // -----------------------------------------------------------------
            string workbookPath = "input.xlsx";
            Workbook workbook = new Workbook(workbookPath);

            // -----------------------------------------------------------------
            // 2. Prepare SVG rendering options
            // -----------------------------------------------------------------
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                // Ensure the generated SVG fits the viewport of the web viewer
                FitToViewPort = true
            };

            // -----------------------------------------------------------------
            // 3. Attach the custom draw‑object handler to capture element bounds
            // -----------------------------------------------------------------
            SvgDrawObjectHandler handler = new SvgDrawObjectHandler();
            svgOptions.DrawObjectEventHandler = handler;

            // -----------------------------------------------------------------
            // 4. Render the first worksheet to SVG (in memory)
            // -----------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            SheetRender renderer = new SheetRender(sheet, svgOptions);

            using (MemoryStream svgStream = new MemoryStream())
            {
                // Render page 0 (the only page because OnePagePerSheet is default for SVG)
                renderer.ToImage(0, svgStream);

                // Save the SVG file for the web viewer
                File.WriteAllBytes("worksheet.svg", svgStream.ToArray());
            }

            // -----------------------------------------------------------------
            // 5. Serialize the captured draw‑object information to JSON
            //    This JSON can be consumed by client‑side JavaScript to perform
            //    hit‑testing (e.g., mapping mouse coordinates to object IDs).
            // -----------------------------------------------------------------
            string json = JsonSerializer.Serialize(handler.ObjectsInfo, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText("drawObjectsMap.json", json);

            Console.WriteLine("SVG and hit‑test map generated successfully.");
        }
    }
}
