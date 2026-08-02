using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ApplyCustomThemeToChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet ws = workbook.Worksheets[0];

        // Populate data for a multi‑series chart
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["A2"].PutValue("Jan");
        ws.Cells["A3"].PutValue("Feb");
        ws.Cells["A4"].PutValue("Mar");

        ws.Cells["B1"].PutValue("Series1");
        ws.Cells["B2"].PutValue(10);
        ws.Cells["B3"].PutValue(20);
        ws.Cells["B4"].PutValue(30);

        ws.Cells["C1"].PutValue("Series2");
        ws.Cells["C2"].PutValue(15);
        ws.Cells["C3"].PutValue(25);
        ws.Cells["C4"].PutValue(35);

        // Add a column chart
        int chartIdx = ws.Charts.Add(ChartType.Column, 6, 0, 20, 10);
        Chart chart = ws.Charts[chartIdx];

        // Set the data range for the series
        chart.NSeries.Add("B1:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Define a custom theme (12 colors as required by CustomTheme)
        Color[] customColors = new Color[]
        {
            Color.White,               // Background1
            Color.Black,               // Text1
            Color.LightGray,           // Background2
            Color.DarkGray,            // Text2
            Color.FromArgb(255, 99, 71),   // Accent1 (Tomato)
            Color.FromArgb(60, 179, 113),  // Accent2 (MediumSeaGreen)
            Color.FromArgb(30, 144, 255),  // Accent3 (DodgerBlue)
            Color.FromArgb(218, 112, 214), // Accent4 (Orchid)
            Color.FromArgb(255, 215, 0),   // Accent5 (Gold)
            Color.FromArgb(255, 140, 0),   // Accent6 (DarkOrange)
            Color.Blue,               // Hyperlink
            Color.Purple              // Followed Hyperlink
        };

        // Apply the custom theme to the workbook
        workbook.CustomTheme("MyCustomTheme", customColors);

        // Apply a monochromatic palette that uses the custom Accent1 color to all series
        SeriesCollection seriesColl = chart.NSeries;
        seriesColl.ChangeColors(ChartColorPaletteType.MonochromaticPalette1);

        // Save the workbook
        workbook.Save("ChartWithCustomTheme.xlsx");
    }
}