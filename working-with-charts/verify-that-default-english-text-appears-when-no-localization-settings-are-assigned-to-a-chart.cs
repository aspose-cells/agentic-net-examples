// Title: Validate that a newly added Aspose.Cells .NET chart displays English default titles and axis labels when no localization is configured
// AI Prompts: Write a C# Aspose.Cells snippet that adds a column chart to a workbook and asserts that the chart title, category axis title, and value axis title match the built‑in English defaults. | Show how to programmatically read and compare a chart’s Title.Text, CategoryAxis.Title.Text, and ValueAxis.Title.Text to the expected English strings in Aspose.Cells without applying any locale settings. | Create an example that saves the workbook after confirming the default English text of all chart elements, and prints a message indicating success or failure.
// Common Searches: aspnet chart default title English Aspose.Cells without localization | c# Aspose.Cells verify default axis labels on a new chart | how to check built‑in chart text in Aspose.Cells .NET | Aspose.Cells chart localization default strings test | default chart title and axis names Aspose.Cells example
// Tags: chart default title verification Aspose.Cells | Aspose.Cells chart axis title English default | C# create column chart Aspose.Cells | Aspose.Cells workbook save after chart validation | chart localization check Aspose.Cells .NET | Aspose.Cells default chart text assertion

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The program creates a workbook, fills sample data, adds a column chart, and confirms that the chart title, category axis title, and value axis title are the built‑in English defaults ('Chart Title', 'Category Axis', 'Value Axis') before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data series for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Verify default chart title (English)
            string defaultTitle = chart.Title.Text;
            if (defaultTitle == "Chart Title")
                Console.WriteLine("Default chart title is English as expected.");
            else
                Console.WriteLine($"Unexpected chart title: '{defaultTitle}'");

            // Verify default category axis title (English)
            string defaultCategoryAxisTitle = chart.CategoryAxis.Title.Text;
            if (defaultCategoryAxisTitle == "Category Axis")
                Console.WriteLine("Default category axis title is English as expected.");
            else
                Console.WriteLine($"Unexpected category axis title: '{defaultCategoryAxisTitle}'");

            // Verify default value axis title (English)
            string defaultValueAxisTitle = chart.ValueAxis.Title.Text;
            if (defaultValueAxisTitle == "Value Axis")
                Console.WriteLine("Default value axis title is English as expected.");
            else
                Console.WriteLine($"Unexpected value axis title: '{defaultValueAxisTitle}'");

            // Save the workbook
            workbook.Save("ChartLocalizationCheck.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
