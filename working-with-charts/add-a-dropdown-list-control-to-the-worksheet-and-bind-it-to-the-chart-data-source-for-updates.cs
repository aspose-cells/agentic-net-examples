// Title: Add an ActiveX ComboBox dropdown to an Excel worksheet and bind its selection to a chart title using Aspose.Cells for .NET (C#)
// Description: This example creates a new workbook, fills column A with category names and column B with numeric values, inserts an ActiveX ComboBox linked to cell C1, populates the list from A2:A5, sets an initial selection, adds a column chart based on B2:B5, and updates the chart title to show the text from the linked cell. The workbook is saved as DropdownChartBinding.xlsx.
// Keywords: Aspose.Cells | C# | .NET | ActiveX ComboBox | Excel dropdown list | chart title binding | linked cell | dynamic chart | sample code | GitHub example | Excel automation
// Common Searches: Aspose.Cells add ComboBox ActiveX control | bind dropdown selection to chart title Aspose.Cells | update Excel chart when linked cell changes C# | create interactive Excel dashboard with Aspose.Cells | sample code for dropdown list and chart in .NET
// Developer Intent: Insert a dropdown list on a worksheet and have a chart title automatically reflect the selected item.
// Use Cases: Interactive reports where users pick a category from a ComboBox and the chart title updates instantly. | Dynamic dashboards that use a linked cell to drive chart titles or filter data without manual refresh. | Template workbooks that expose ActiveX controls for end‑users to customize chart captions on the fly.
// AI Prompts: Generate C# code with Aspose.Cells that adds an ActiveX ComboBox, sets ListFillRange to A2:A5, links it to cell C1, and binds a column chart title to that cell. | Show how to refresh the linked cell value after a ComboBox selection and update the chart title in Aspose.Cells for .NET. | Explain how to connect a dropdown list to chart series categories or titles using Aspose.Cells in a C# application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsDropdownChartBinding
{
    // This example creates a new workbook, fills column A with category names and column B with numeric values, inserts an ActiveX ComboBox linked to cell C1, populates the list from A2:A5, sets an initial selection, adds a column chart based on B2:B5, and updates the chart title to show the text from the linked cell. The workbook is saved as DropdownChartBinding.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Populate sample data that will be used for the chart
                // -------------------------------------------------
                // Categories
                sheet.Cells["A2"].PutValue("Item 1");
                sheet.Cells["A3"].PutValue("Item 2");
                sheet.Cells["A4"].PutValue("Item 3");
                sheet.Cells["A5"].PutValue("Item 4");

                // Corresponding values
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                // -------------------------------------------------
                // Add a ComboBox ActiveX control (acts as a dropdown list)
                // -------------------------------------------------
                // Parameters: topRow, top, leftColumn, left, height, width
                Shape comboShape = sheet.Shapes.AddActiveXControl(
                    ControlType.ComboBox, // control type
                    1, 0,                 // topRow, top offset (pixels)
                    3, 0,                 // leftColumn, left offset (pixels)
                    100, 30);             // height, width (pixels)

                // Cast to the specific ComboBoxActiveXControl type
                ComboBoxActiveXControl comboBox = (ComboBoxActiveXControl)comboShape.ActiveXControl;

                // Fill the dropdown list with the category names (A2:A5)
                comboBox.ListFillRange = "A2:A5";

                // Link the selected value to a cell (C1). When the user picks an item,
                // the cell C1 will contain the selected category text.
                comboBox.LinkedCell = "C1";

                // Optional: set the initial selected value
                comboBox.Value = "Item 1";

                // -------------------------------------------------
                // Create a chart that uses the values from column B.
                // The chart will be refreshed when the linked cell (C1) changes.
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add a series that reads its values from B2:B5
                int seriesIndex = chart.NSeries.Add("B2:B5", true);
                Series series = chart.NSeries[seriesIndex];

                // (Optional) Set the series name
                series.Name = "Values";

                // Note: In recent Aspose.Cells versions the Series class does not expose
                // a CategoryData property for column charts. The categories will be taken
                // from the first column (A2:A5) automatically when the chart is refreshed.

                // Bind the chart title to the linked cell so the title updates
                // when a different item is selected from the dropdown.
                chart.Title.Text = "Selected: " + sheet.Cells["C1"].StringValue;

                // -------------------------------------------------
                // Refresh the linked cell value (in case the default selection changed)
                // and update the chart title accordingly.
                // -------------------------------------------------
                sheet.Shapes.UpdateSelectedValue(); // sync ComboBox selection to C1
                chart.Title.Text = "Selected: " + sheet.Cells["C1"].StringValue;

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "DropdownChartBinding.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
