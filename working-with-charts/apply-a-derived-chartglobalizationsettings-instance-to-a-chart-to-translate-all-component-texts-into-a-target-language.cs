using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartGlobalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set a placeholder title (will be overridden by globalization settings)
            chart.Title.Text = "Demo Chart";

            // Apply custom chart globalization settings (e.g., Spanish translation)
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new MyChartGlobalizationSettings()
            };

            // Save the workbook
            workbook.Save("ChartGlobalizationDemo.xlsx");
        }
    }

    // Custom ChartGlobalizationSettings implementation for target language translation
    class MyChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetChartTitleName()
        {
            return "Título del Gráfico";
        }

        public override string GetSeriesName()
        {
            return "Serie";
        }

        public override string GetLegendIncreaseName()
        {
            return "Aumento";
        }

        public override string GetLegendDecreaseName()
        {
            return "Disminución";
        }

        public override string GetLegendTotalName()
        {
            return "Total";
        }

        public override string GetOtherName()
        {
            return "Otro";
        }

        public override string GetAxisTitleName()
        {
            return "Título del Eje";
        }

        public override string GetAxisUnitName(DisplayUnitType type)
        {
            switch (type)
            {
                case DisplayUnitType.Hundreds:
                    return "Cientos";
                case DisplayUnitType.Thousands:
                    return "Miles";
                case DisplayUnitType.TenThousands:
                    return "Diez Miles";
                default:
                    return base.GetAxisUnitName(type);
            }
        }
    }
}