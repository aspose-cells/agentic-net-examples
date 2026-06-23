using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgRotateDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Year");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue(2019);
            sheet.Cells["A3"].PutValue(2020);
            sheet.Cells["A4"].PutValue(2021);
            sheet.Cells["A5"].PutValue(2022);
            sheet.Cells["B2"].PutValue(4500);
            sheet.Cells["B3"].PutValue(4900);
            sheet.Cells["B4"].PutValue(5200);
            sheet.Cells["B5"].PutValue(6100);

            // Add a line chart
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.SetChartDataRange("A1:B5", true);

            // Export the chart to SVG
            string svgPath = "chart.svg";
            SvgImageOptions svgOptions = new SvgImageOptions(); // default options
            chart.ToImage(svgPath, svgOptions);

            // Load the generated SVG content
            string svgContent = File.ReadAllText(svgPath);

            // Insert a CSS rotate transform (45 degrees) into the root <svg> element
            // We add a style attribute: style="transform: rotate(45deg);"
            const string svgTag = "<svg";
            int svgTagPos = svgContent.IndexOf(svgTag, StringComparison.OrdinalIgnoreCase);
            if (svgTagPos >= 0)
            {
                int insertPos = svgContent.IndexOf('>', svgTagPos);
                if (insertPos > 0)
                {
                    // Build the style attribute
                    string styleAttribute = " style=\"transform: rotate(45deg);\"";
                    // Insert the style attribute just before the closing '>'
                    svgContent = svgContent.Insert(insertPos, styleAttribute);
                }
            }

            // Save the modified SVG back to file
            File.WriteAllText(svgPath, svgContent);

            // Optionally, save the workbook (not required for SVG output)
            workbook.Save("ChartWorkbook.xlsx");

            Console.WriteLine("SVG chart saved with 45-degree CSS rotation at: " + Path.GetFullPath(svgPath));
        }
    }
}