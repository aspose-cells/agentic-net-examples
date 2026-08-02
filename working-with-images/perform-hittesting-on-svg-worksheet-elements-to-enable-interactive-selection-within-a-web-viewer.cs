// Title: Create Interactive SVG from Excel with Hit‑Test Data using Aspose.Cells for .NET
// Description: A C# example that builds a workbook, adds a rectangle shape, renders the sheet to SVG, and uses a custom DrawObjectEventHandler to capture each object's type, sheet index, page number and bounding box. The hit‑test information is saved as JSON, enabling client‑side click detection and interactivity in web viewers.
// Keywords: Aspose.Cells SVG | C# hit test | DrawObjectEventHandler | Excel to SVG clickable | SVG element coordinates | export worksheet as SVG | interactive SVG Excel | JSON hit‑test mapping | Aspose.Cells .NET example
// Common Searches: Aspose.Cells capture SVG element positions | C# render Excel worksheet to SVG with hit test | How to get bounding boxes of shapes in SVG using Aspose.Cells | Export Excel to SVG and generate click map | DrawObjectEventHandler example for SVG
// Developer Intent: Collect geometric data of every drawn object during SVG export so the front‑end can identify which element a user clicked.
// Use Cases: Generate an SVG view of a spreadsheet and overlay JavaScript click handlers based on JSON‑encoded bounding boxes. | Create tooltips or pop‑ups for charts, images, and shapes by mapping SVG coordinates back to the original workbook. | Implement a web‑based spreadsheet viewer that allows users to select or edit objects directly on the SVG canvas.
// AI Prompts: Extend SvgHitTestHandler to include object IDs and text content in the JSON output. | Show JavaScript code that loads HitTestMapping.json and highlights the clicked SVG shape. | Explain how to configure SvgImageOptions to fit the SVG to the viewport while preserving hit‑test data.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Custom handler to capture draw object details during SVG rendering
// A C# example that builds a workbook, adds a rectangle shape, renders the sheet to SVG, and uses a custom DrawObjectEventHandler to capture each object's type, sheet index, page number and bounding box. The hit‑test information is saved as JSON, enabling client‑side click detection and interactivity in web viewers.
class SvgHitTestHandler : DrawObjectEventHandler
{
    // List to store hit‑test information for each draw object
    public List<Dictionary<string, object>> HitTestData { get; } = new List<Dictionary<string, object>>();

    // Called for every draw object (shapes, charts, images, etc.)
    public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
    {
        var info = new Dictionary<string, object>
        {
            { "Type", drawObject.Type.ToString() },
            { "SheetIndex", drawObject.SheetIndex },
            { "Page", drawObject.CurrentPage + 1 }, // 1‑based page number
            { "TotalPages", drawObject.TotalPages },
            { "X", x },
            { "Y", y },
            { "Width", width },
            { "Height", height }
        };

        HitTestData.Add(info);
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // -------------------------------------------------
            // 1. Create a workbook and populate sample data
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Demo";

            // Sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(85);

            // Add a rectangle shape that we want to be clickable in the SVG
            Shape rect = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 200, 80);
            rect.Text = "Click Me";

            // -------------------------------------------------
            // 2. Prepare SVG rendering options
            // -------------------------------------------------
            ImageOrPrintOptions renderOpts = new ImageOrPrintOptions
            {
                ImageType = ImageType.Svg,
                OnePagePerSheet = true
                // FitToViewPort can be set via SvgImageOptions if the property is available in the used version
            };

            // Attach the custom draw‑object handler to collect hit‑test data
            SvgHitTestHandler hitTestHandler = new SvgHitTestHandler();
            renderOpts.DrawObjectEventHandler = hitTestHandler;

            // -------------------------------------------------
            // 3. Render the worksheet to an SVG file
            // -------------------------------------------------
            SheetRender renderer = new SheetRender(sheet, renderOpts);
            string svgPath = "WorksheetWithHitTest.svg";
            renderer.ToImage(0, svgPath);

            // -------------------------------------------------
            // 4. Persist the hit‑test mapping as a JSON file
            // -------------------------------------------------
            string jsonPath = "HitTestMapping.json";
            string json = JsonSerializer.Serialize(hitTestHandler.HitTestData, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(jsonPath, json);

            // -------------------------------------------------
            // 5. Save the original workbook (optional)
            // -------------------------------------------------
            string workbookPath = "DemoWorkbook.xlsx";
            workbook.Save(workbookPath);

            Console.WriteLine($"SVG generated: {svgPath}");
            Console.WriteLine($"Hit‑test data saved: {jsonPath}");
            Console.WriteLine($"Workbook saved: {workbookPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
