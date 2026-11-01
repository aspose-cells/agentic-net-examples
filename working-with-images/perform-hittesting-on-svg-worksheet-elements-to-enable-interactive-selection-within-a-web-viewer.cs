// Title: How to add hit‑testing to SVG elements generated from an Excel worksheet using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells render worksheet to SVG and detect which shape was clicked | Parse SVG rect and text nodes from Aspose.Cells output for hit testing | Implement point‑in‑rectangle hit test on SVG elements produced by Aspose.Cells | Interactive selection of Excel cells in a web viewer using SVG hit detection in C# | Extract bounding boxes from Aspose.Cells SVG and perform click hit test
// Tags: Aspose.Cells SVG hit testing | C# parse SVG rect nodes | SVG text bounding box approximation C# | SheetRender SVG element extraction | interactive Excel worksheet selection SVG

using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgHitTest
{
    // Represents a simple SVG element with a rectangular bounding box.
    // The example loads an Excel workbook, uses Aspose.Cells SheetRender to convert the first worksheet to SVG, parses <rect> and <text> elements to build bounding‑box objects, stores them in a list, and provides a HitTest method that returns the element ID or type when a given SVG coordinate falls inside any element's bounds.
    public class SvgElementInfo
    {
        public string Id { get; set; } = string.Empty;   // Optional identifier (e.g., shape name)
        public string Type { get; set; } = string.Empty; // e.g., "rect", "text"
        public double X { get; set; }                    // Upper‑left X
        public double Y { get; set; }                    // Upper‑left Y
        public double Width { get; set; }
        public double Height { get; set; }

        // Determines whether a point (px,py) lies inside the element's bounding box.
        public bool Contains(double px, double py)
        {
            return px >= X && px <= X + Width && py >= Y && py <= Y + Height;
        }
    }

    public class SvgHitTester
    {
        private readonly List<SvgElementInfo> _elements = new List<SvgElementInfo>();

        // Load workbook, render the first worksheet to SVG and parse hit‑testable elements.
        public void LoadAndParse(string workbookPath)
        {
            try
            {
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                    return;
                }

                // Load workbook
                Workbook wb = new Workbook(workbookPath);

                // Render first worksheet to SVG
                Worksheet sheet = wb.Worksheets[0];
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    SaveFormat = SaveFormat.Svg,   // Correct property for SVG output
                    OnePagePerSheet = true
                };
                SheetRender sr = new SheetRender(sheet, options);

                using (MemoryStream svgStream = new MemoryStream())
                {
                    sr.ToImage(0, svgStream);
                    svgStream.Position = 0;
                    ParseSvg(svgStream);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception during LoadAndParse: {ex.Message}");
            }
        }

        // Parse SVG stream and extract rectangular elements (rect, text with bounding box approximations).
        private void ParseSvg(Stream svgStream)
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(svgStream);
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("svg", "http://www.w3.org/2000/svg");

                // Parse <rect> elements
                XmlNodeList? rectNodes = doc.SelectNodes("//svg:rect", nsmgr);
                if (rectNodes != null)
                {
                    foreach (XmlNode node in rectNodes)
                    {
                        var elem = new SvgElementInfo
                        {
                            Id = node.Attributes?["id"]?.Value ?? string.Empty,
                            Type = "rect",
                            X = ParseDouble(node.Attributes?["x"]?.Value),
                            Y = ParseDouble(node.Attributes?["y"]?.Value),
                            Width = ParseDouble(node.Attributes?["width"]?.Value),
                            Height = ParseDouble(node.Attributes?["height"]?.Value)
                        };
                        _elements.Add(elem);
                    }
                }

                // Parse <text> elements – approximate bounding box using font size (if present)
                XmlNodeList? textNodes = doc.SelectNodes("//svg:text", nsmgr);
                if (textNodes != null)
                {
                    foreach (XmlNode node in textNodes)
                    {
                        double x = ParseDouble(node.Attributes?["x"]?.Value);
                        double y = ParseDouble(node.Attributes?["y"]?.Value);
                        double fontSize = 12; // default fallback
                        if (node.Attributes?["font-size"] != null)
                            fontSize = ParseDouble(node.Attributes?["font-size"]?.Value);
                        string text = node.InnerText ?? string.Empty;
                        double width = text.Length * fontSize * 0.6;
                        double height = fontSize;

                        var elem = new SvgElementInfo
                        {
                            Id = node.Attributes?["id"]?.Value ?? string.Empty,
                            Type = "text",
                            X = x,
                            Y = y - height, // SVG y refers to baseline; shift to top
                            Width = width,
                            Height = height
                        };
                        _elements.Add(elem);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception during ParseSvg: {ex.Message}");
            }
        }

        // Helper to safely parse double values; returns 0 on failure.
        private static double ParseDouble(string? s)
        {
            if (double.TryParse(s, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double result))
                return result;
            return 0;
        }

        // Perform hit‑testing given a point in SVG coordinate space.
        // Returns the Id of the first element that contains the point, or null if none.
        public string? HitTest(double svgX, double svgY)
        {
            foreach (var elem in _elements)
            {
                if (elem.Contains(svgX, svgY))
                    return string.IsNullOrEmpty(elem.Id) ? elem.Type : elem.Id;
            }
            return null;
        }
    }

    // Entry point demonstrating the workflow.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file (use provided load rule if applicable)
            string excelPath = "input.xlsx";

            SvgHitTester tester = new SvgHitTester();
            tester.LoadAndParse(excelPath);

            // Simulated click coordinates (replace with actual client coordinates transformed to SVG space)
            double clickX = 120.0;
            double clickY = 80.0;

            string? hitResult = tester.HitTest(clickX, clickY);
            Console.WriteLine(hitResult != null
                ? $"Hit element: {hitResult}"
                : "No SVG element was hit at the given coordinates.");
        }
    }
}
