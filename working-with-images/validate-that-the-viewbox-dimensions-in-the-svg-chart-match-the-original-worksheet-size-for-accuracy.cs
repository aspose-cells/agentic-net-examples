using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

class SvgViewBoxValidator
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart covering rows 5-20 and columns 0-8
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Obtain the actual size of the chart in pixels (width, height)
        int[] actualSize = chart.GetActualSize(); // uses Chart.GetActualSize rule
        int chartWidthPx = actualSize[0];
        int chartHeightPx = actualSize[1];

        // Configure SVG rendering options to fit to viewport
        SvgImageOptions svgOptions = new SvgImageOptions();
        svgOptions.FitToViewPort = true; // uses SvgImageOptions.FitToViewPort rule

        // Render the first worksheet to an SVG file
        SheetRender renderer = new SheetRender(worksheet, svgOptions);
        string svgPath = "chart_output.svg";
        renderer.ToImage(0, svgPath); // saves the SVG

        // Load the generated SVG and extract the viewBox attribute
        XDocument svgDoc = XDocument.Load(svgPath);
        XAttribute viewBoxAttr = svgDoc.Root.Attribute("viewBox");
        if (viewBoxAttr == null)
        {
            Console.WriteLine("The SVG does not contain a viewBox attribute.");
            return;
        }

        // viewBox format: "minX minY width height"
        string[] viewBoxParts = viewBoxAttr.Value.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
        if (viewBoxParts.Length != 4)
        {
            Console.WriteLine("Unexpected viewBox format.");
            return;
        }

        // Parse width and height from viewBox (they are the 3rd and 4th values)
        double viewBoxWidth = double.Parse(viewBoxParts[2]);
        double viewBoxHeight = double.Parse(viewBoxParts[3]);

        // Compare SVG viewBox dimensions with the chart's actual pixel size
        const double tolerance = 0.5; // allow a small rounding tolerance
        bool widthMatches = Math.Abs(viewBoxWidth - chartWidthPx) <= tolerance;
        bool heightMatches = Math.Abs(viewBoxHeight - chartHeightPx) <= tolerance;

        Console.WriteLine($"Chart actual size:   Width = {chartWidthPx}px, Height = {chartHeightPx}px");
        Console.WriteLine($"SVG viewBox size:    Width = {viewBoxWidth}px, Height = {viewBoxHeight}px");
        Console.WriteLine($"Width match:  {widthMatches}");
        Console.WriteLine($"Height match: {heightMatches}");

        // Optionally, save the workbook (demonstrates lifecycle save rule)
        workbook.Save("validated_chart.xlsx");
    }
}