using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class Remove3DFormatting
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a 3‑D column chart with sample data
        worksheet.Cells["A1"].PutValue("Category 1");
        worksheet.Cells["A2"].PutValue("Category 2");
        worksheet.Cells["B1"].PutValue(10);
        worksheet.Cells["B2"].PutValue(20);

        int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 15, 10);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B1:B2", true);
        chart.NSeries.CategoryData = "A1:A2";

        // OPTIONAL: Apply some 3‑D formatting to the first series (to demonstrate the clearing)
        Series series = chart.NSeries[0];
        ShapePropertyCollection shapeProps = series.ShapeProperties;
        Format3D format3D = shapeProps.Format3D;
        format3D.TopBevel.Type = BevelPresetType.Circle;
        format3D.TopBevel.Height = 2;
        format3D.TopBevel.Width = 5;
        format3D.SurfaceMaterialType = PresetMaterialType.WarmMatte;
        format3D.SurfaceLightingType = LightRigType.ThreePoint;
        format3D.LightingAngle = 20;

        // Iterate through all series in the chart and clear any 3‑D formatting
        foreach (Series s in chart.NSeries)
        {
            ShapePropertyCollection spc = s.ShapeProperties;
            if (spc.HasFormat3D())
            {
                spc.ClearFormat3D(); // Removes all 3‑D properties, preserving 2‑D settings
            }
        }

        // Save the workbook
        workbook.Save("Removed3DFormatting.xlsx", SaveFormat.Xlsx);
    }
}