// Title: Create a dropdown list in Excel and bind it to a column chart series using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that adds a data‑validation dropdown in cell D1 containing "B" and "C", then sets a column chart’s Y‑values to an INDIRECT formula that references the selected column. | Show how to generate a dynamic column chart whose series data updates automatically when the worksheet dropdown selection changes, using the Aspose.Cells .NET API. | Provide a complete Aspose.Cells example that creates sample data, inserts the dropdown, configures the chart binding, and saves the workbook as DropdownChart.xlsx.
// Common Searches: how to link Excel dropdown to chart series using Aspose.Cells C# | Aspose.Cells create data validation dropdown that controls chart data source | dynamic chart values with INDIRECT formula in Aspose.Cells .NET | C# example for updating column chart based on cell selection in Excel
// Tags: Aspose.Cells dropdown validation | Aspose.Cells bind chart series to cell value | Aspose.Cells INDIRECT formula dynamic chart | C# column chart selectable series | Excel workbook dropdown driven chart update

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The program creates a workbook, adds sample data, inserts a data‑validation dropdown in D1 offering columns B and C, creates a column chart, sets the chart's Y‑values to an INDIRECT formula that references the column chosen in the dropdown, and saves the file as DropdownChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data for two series in columns B and C
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["A2"].PutValue(10);
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["C2"].PutValue(20);
                sheet.Cells["A3"].PutValue(20);
                sheet.Cells["B3"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["A4"].PutValue(30);
                sheet.Cells["B4"].PutValue(20);
                sheet.Cells["C4"].PutValue(30);
                sheet.Cells["A5"].PutValue(40);
                sheet.Cells["B5"].PutValue(25);
                sheet.Cells["C5"].PutValue(35);

                // Add a dropdown list (data validation) in cell D1 to select the data column
                CellArea validationArea = CellArea.CreateCellArea("D1", "D1");
                int validationIndex = sheet.Validations.Add(validationArea);
                Validation validation = sheet.Validations[validationIndex];
                validation.Type = ValidationType.List;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = "\"B,C\""; // List items: column letters B and C

                // Show input message (prompt) for the validation
                validation.InputTitle = "Select Series";
                validation.InputMessage = "Choose the column to display in the chart.";
                // Note: ShowInputMessage property is not required; default behavior displays the prompt.

                // Create a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // X‑axis categories from A2:A5
                chart.NSeries.Add("=Sheet1!$A$2:$A$5", true);

                // Y‑axis values use INDIRECT to refer to the column chosen in the dropdown (D1)
                chart.NSeries[0].Values = "=INDIRECT(\"Sheet1!\" & $D$1 & \"$2:$\" & $D$1 & \"$5\")";

                // Optional: set chart title
                chart.Title.Text = "Dynamic Series Based on Dropdown";

                // Save the workbook
                workbook.Save("DropdownChart.xlsx");
                Console.WriteLine("Workbook saved successfully as 'DropdownChart.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
