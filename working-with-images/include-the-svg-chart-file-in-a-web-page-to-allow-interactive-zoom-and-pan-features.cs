// Title: Generate an SVG column chart with Aspose.Cells in C# and embed it in an HTML page with client‑side zoom and pan
// Common Searches: C# Aspose.Cells export column chart to SVG file | How to embed an Aspose.Cells generated SVG chart in a web page | JavaScript mouse wheel zoom and drag pan for an embedded SVG chart | Render Excel chart as SVG using Aspose.Cells and create an interactive viewer | Aspose.Cells ImageOrPrintOptions settings for SVG output in web applications
// Tags: Aspose.Cells export chart to SVG | C# SheetRender ToImage SVG output | ImageOrPrintOptions SVG rendering options | JavaScript viewBox zoom pan for embedded SVG | HTML object element embed SVG chart | Aspose.Cells column chart generation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsSvgChartWeb
{
    // The sample creates a workbook with fruit‑sales data, adds a column chart, and uses Aspose.Cells ImageOrPrintOptions and SheetRender to export the chart as an SVG file. It then generates an HTML page that loads the SVG via an <object> element and implements simple JavaScript that adjusts the SVG viewBox to provide mouse‑wheel zoom and drag‑to‑pan interactivity.
    class Program
    {
        static void Main(string[] args)
        {
            // Define output file names (relative to the executable folder)
            const string svgFileName = "chart.svg";
            const string htmlFileName = "chart.html";

            try
            {
                // ------------------------------------------------------------
                // 1. Create a workbook and populate it with sample data
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Header
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // Sample rows
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(800);
                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["B4"].PutValue(1500);

                // ------------------------------------------------------------
                // 2. Add a column chart that uses the data above
                // ------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";
                chart.Title.Text = "Fruit Sales";

                // ------------------------------------------------------------
                // 3. Configure rendering options to produce an SVG file
                // ------------------------------------------------------------
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Svg,          // Output format
                    OnePagePerSheet = true,             // Render the whole sheet as one SVG page
                    SvgCssPrefix = "chart_",            // CSS prefix to avoid style collisions (obsolete but still functional)
                    SVGFitToViewPort = true             // Fit SVG to viewport (obsolete but still functional)
                };

                // ------------------------------------------------------------
                // 4. Render the worksheet (which contains the chart) to SVG
                // ------------------------------------------------------------
                SheetRender sheetRender = new SheetRender(sheet, renderOptions);
                // Ensure the output directory exists
                string svgFullPath = Path.GetFullPath(svgFileName);
                Directory.CreateDirectory(Path.GetDirectoryName(svgFullPath) ?? string.Empty);
                sheetRender.ToImage(0, svgFullPath);
                Console.WriteLine($"SVG chart saved to: {svgFullPath}");

                // ------------------------------------------------------------
                // 5. Build an HTML page that loads the SVG and adds zoom/pan
                // ------------------------------------------------------------
                string htmlContent = $@"
<!DOCTYPE html>
<html lang=""en"">
<head>
    <meta charset=""UTF-8"">
    <title>Aspose.Cells SVG Chart with Zoom & Pan</title>
    <style>
        body {{ margin:0; overflow:hidden; }}
        #svgContainer {{
            width:100vw;
            height:100vh;
            background:#f0f0f0;
            display:flex;
            justify-content:center;
            align-items:center;
        }}
        svg {{ border:1px solid #ccc; }}
    </style>
</head>
<body>
    <div id=""svgContainer"">
        <!-- Embed the generated SVG -->
        <object id=""chartSvg"" type=""image/svg+xml"" data=""{svgFileName}"" width=""90%"" height=""90%""></object>
    </div>

    <script>
        // Simple zoom & pan implementation for the embedded SVG
        (function () {{
            const obj = document.getElementById('chartSvg');
            obj.addEventListener('load', function () {{
                const svgDoc = obj.contentDocument;
                const svgRoot = svgDoc.documentElement;

                // Ensure viewBox exists
                let viewBox = svgRoot.getAttribute('viewBox');
                if (!viewBox) {{
                    const w = svgRoot.getAttribute('width') || 800;
                    const h = svgRoot.getAttribute('height') || 600;
                    viewBox = `0 0 ${{w}} ${{h}}`;
                    svgRoot.setAttribute('viewBox', viewBox);
                }}

                let [vbX, vbY, vbW, vbH] = viewBox.split(' ').map(Number);
                const scaleFactor = 0.1; // Zoom step (10%)
                let isPanning = false;
                let startX, startY;

                // Zoom with mouse wheel
                obj.addEventListener('wheel', function (e) {{
                    e.preventDefault();
                    const delta = Math.sign(e.deltaY);
                    const zoom = delta > 0 ? (1 - scaleFactor) : (1 + scaleFactor);
                    const mx = e.offsetX / obj.clientWidth;
                    const my = e.offsetY / obj.clientHeight;

                    const newW = vbW * zoom;
                    const newH = vbH * zoom;
                    const newX = vbX + (vbW - newW) * mx;
                    const newY = vbY + (vbH - newH) * my;

                    vbX = newX; vbY = newY; vbW = newW; vbH = newH;
                    svgRoot.setAttribute('viewBox', `${{vbX}} ${{vbY}} ${{vbW}} ${{vbH}}`);
                }});

                // Pan with mouse drag
                obj.addEventListener('mousedown', function (e) {{
                    if (e.button !== 0) return;
                    isPanning = true;
                    startX = e.clientX;
                    startY = e.clientY;
                }});
                window.addEventListener('mousemove', function (e) {{
                    if (!isPanning) return;
                    const dx = (e.clientX - startX) * (vbW / obj.clientWidth);
                    const dy = (e.clientY - startY) * (vbH / obj.clientHeight);
                    vbX -= dx;
                    vbY -= dy;
                    startX = e.clientX;
                    startY = e.clientY;
                    svgRoot.setAttribute('viewBox', `${{vbX}} ${{vbY}} ${{vbW}} ${{vbH}}`);
                }});
                window.addEventListener('mouseup', function () {{
                    isPanning = false;
                }});
            }});
        }})();
    </script>
</body>
</html>";

                // Write the HTML file
                string htmlFullPath = Path.GetFullPath(htmlFileName);
                Directory.CreateDirectory(Path.GetDirectoryName(htmlFullPath) ?? string.Empty);
                File.WriteAllText(htmlFullPath, htmlContent);
                Console.WriteLine($"HTML page saved to: {htmlFullPath}");
                Console.WriteLine("Open the HTML file in a browser to view the interactive SVG chart.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred while generating the SVG chart or HTML page:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
