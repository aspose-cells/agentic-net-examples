using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Export the chart to SVG
        SvgImageOptions svgOptions = new SvgImageOptions();
        string originalSvgPath = "chart.svg";
        chart.ToImage(originalSvgPath, svgOptions);

        // Load the generated SVG content
        string svgContent = File.ReadAllText(originalSvgPath, Encoding.UTF8);

        // Insert a CSS rotate transform (45 degrees) into the root <svg> element
        int svgTagStart = svgContent.IndexOf("<svg");
        if (svgTagStart >= 0)
        {
            int svgTagEnd = svgContent.IndexOf('>', svgTagStart);
            if (svgTagEnd > svgTagStart)
            {
                string beforeTag = svgContent.Substring(0, svgTagEnd);
                string afterTag = svgContent.Substring(svgTagEnd);
                // Add style attribute only if it does not already exist
                if (!beforeTag.Contains("style=\""))
                {
                    string styleAttribute = " style=\"transform: rotate(45deg);\"";
                    svgContent = beforeTag + styleAttribute + afterTag;
                }
            }
        }

        // Save the modified SVG with rotation applied
        string rotatedSvgPath = "chart_rotated.svg";
        File.WriteAllText(rotatedSvgPath, svgContent, Encoding.UTF8);

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("output.xlsx");
    }
}