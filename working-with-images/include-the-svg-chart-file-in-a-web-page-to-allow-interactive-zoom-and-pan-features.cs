using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgWebDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a workbook and populate it with sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");

                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(12000);
                sheet.Cells["B3"].PutValue(15000);
                sheet.Cells["B4"].PutValue(18000);
                sheet.Cells["B5"].PutValue(21000);

                // 2. Add a line chart that uses the data
                int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B5", true);          // Values
                chart.NSeries.CategoryData = "A2:A5";      // Categories
                chart.Title.Text = "Monthly Sales";

                // 3. Configure SVG rendering options – enable FitToViewPort for responsive scaling
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    FitToViewPort = true,
                    CssPrefix = "chart-"
                };

                // 4. Render the chart to an SVG file
                string svgPath = "chart.svg";
                chart.ToImage(svgPath, svgOptions);

                // 5. Load the generated SVG content (ensure file exists)
                if (!File.Exists(svgPath))
                    throw new FileNotFoundException("SVG file was not created.", svgPath);

                string svgContent = File.ReadAllText(svgPath);

                // 6. Build an HTML page that embeds the SVG and adds simple zoom/pan support
                string htmlTemplate = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset=""utf-8"" />
    <title>Interactive SVG Chart</title>
    <style>
        /* Ensure the SVG fills the container */
        #svgContainer {{
            width: 100%;
            height: 80vh;
            border: 1px solid #ccc;
            overflow: hidden;
            position: relative;
        }}
        svg {{
            width: 100%;
            height: 100%;
            cursor: grab;
        }}
    </style>
</head>
<body>
    <h2>Interactive Zoom & Pan for Aspose.Cells SVG Chart</h2>
    <div id=""svgContainer"">
        {svgContent}
    </div>

    <script>
        // Simple zoom/pan implementation using viewBox manipulation
        const svg = document.querySelector('#svgContainer svg');
        let viewBox = svg.getAttribute('viewBox').split(' ').map(Number);
        let isPanning = false;
        let startX, startY;

        // Mouse wheel for zoom
        svg.addEventListener('wheel', function (e) {{
            e.preventDefault();
            const scaleFactor = e.deltaY < 0 ? 0.9 : 1.1; // zoom in/out
            const [x, y, w, h] = viewBox;
            const mx = e.offsetX / svg.clientWidth;
            const my = e.offsetY / svg.clientHeight;
            const newW = w * scaleFactor;
            const newH = h * scaleFactor;
            const newX = x + (w - newW) * mx;
            const newY = y + (h - newH) * my;
            viewBox = [newX, newY, newW, newH];
            svg.setAttribute('viewBox', viewBox.join(' '));
        }});

        // Mouse down to start panning
        svg.addEventListener('mousedown', function (e) {{
            isPanning = true;
            startX = e.clientX;
            startY = e.clientY;
            svg.style.cursor = 'grabbing';
        }});

        // Mouse move to pan
        svg.addEventListener('mousemove', function (e) {{
            if (!isPanning) return;
            const dx = (e.clientX - startX) * (viewBox[2] / svg.clientWidth);
            const dy = (e.clientY - startY) * (viewBox[3] / svg.clientHeight);
            viewBox[0] -= dx;
            viewBox[1] -= dy;
            svg.setAttribute('viewBox', viewBox.join(' '));
            startX = e.clientX;
            startY = e.clientY;
        }});

        // Mouse up to stop panning
        window.addEventListener('mouseup', function () {{
            isPanning = false;
            svg.style.cursor = 'grab';
        }});
    </script>
</body>
</html>";

                // 7. Save the HTML file
                string htmlPath = "chart.html";
                File.WriteAllText(htmlPath, htmlTemplate);

                Console.WriteLine($"SVG chart generated at: {Path.GetFullPath(svgPath)}");
                Console.WriteLine($"Interactive HTML page generated at: {Path.GetFullPath(htmlPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}