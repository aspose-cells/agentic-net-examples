// Title: Create a ComboBox dropdown in an Excel worksheet and bind it to a dynamic chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add ComboBox to worksheet and link it to chart series | How to make Excel chart update when ComboBox selection changes using Aspose.Cells | C# code for binding a dropdown list to a chart data source with Aspose.Cells | Dynamic chart series selection with ComboBox in Aspose.Cells .NET | Example of using InputRange and LinkedCell properties for chart binding in Aspose.Cells
// Tags: add combobox control to Excel worksheet Aspose.Cells | chart series binding via dropdown Aspose.Cells | formula-driven helper column for dynamic chart Aspose.Cells | update Excel chart on combobox change C# | use InputRange and LinkedCell for chart data Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

namespace AsposeCellsDropdownChartBinding
{
    // The sample creates a new workbook, populates category and series data, inserts a ComboBox at D2 whose items are sourced from the series headers, links the selected value to cell D2, builds a helper column with IF formulas that select values from the chosen series, creates a column chart that reads from the helper column, recalculates formulas, and saves the workbook as DropdownBoundChart.xlsx, enabling the chart to update automatically when the dropdown selection changes.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 1. Populate sample data
                // ------------------------------------------------------------
                // Headers
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");

                // Categories
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");

                // Values for Series1
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                // Values for Series2
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);
                sheet.Cells["C5"].PutValue(45);

                // ------------------------------------------------------------
                // 2. Add a ComboBox (dropdown) that lets the user pick a series
                // ------------------------------------------------------------
                // The ComboBox will be placed at D2 (row 1, column 3) with a size of 100x20 points
                ComboBox comboBox = sheet.Shapes.AddComboBox(1, 3, 0, 0, 100, 20);
                // Show a reasonable number of items in the dropdown list
                comboBox.DropDownLines = 5;
                // The list of items comes from the header cells B1:C1
                comboBox.InputRange = "B1:C1";
                // The cell that will receive the selected value
                comboBox.LinkedCell = "D2";
                // Optional: give the ComboBox a name for easier reference
                comboBox.Name = "SeriesSelector";

                // ------------------------------------------------------------
                // 3. Create a helper column that reflects the selected series
                // ------------------------------------------------------------
                // Column E will contain a formula that picks the value from Series1 or Series2
                // based on the text stored in D2 (the linked cell of the ComboBox).
                for (int row = 2; row <= 5; row++)
                {
                    // =IF($D$2="Series1",$B2,IF($D$2="Series2",$C2,0))
                    string formula = $"IF($D$2=\"Series1\",$B{row},IF($D$2=\"Series2\",$C{row},0))";
                    sheet.Cells[$"E{row}"].Formula = formula;
                }

                // Header for the helper column (optional, not used by the chart)
                sheet.Cells["E1"].PutValue("SelectedSeries");

                // ------------------------------------------------------------
                // 4. Add a chart that uses the helper column as its data source
                // ------------------------------------------------------------
                // Place the chart at G2 (row 1, column 6) with a size of 400x300 points
                int chartIndex = sheet.Charts.Add(ChartType.Column, 1, 6, 20, 30);
                Chart chart = sheet.Charts[chartIndex];

                // Set the series data to the helper column (E2:E5)
                chart.NSeries.Add("E2:E5", true);
                // Set the category (X‑axis) labels to the categories (A2:A5)
                chart.NSeries.CategoryData = "A2:A5";

                // Give the chart a title
                chart.Title.Text = "Dynamic Series Chart";

                // ------------------------------------------------------------
                // 5. Recalculate formulas so the chart reflects the initial selection
                // ------------------------------------------------------------
                workbook.CalculateFormula();

                // ------------------------------------------------------------
                // 6. Save the workbook
                // ------------------------------------------------------------
                string outputPath = "DropdownBoundChart.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to {outputPath}");
                Console.WriteLine("Open the file, change the dropdown (cell D2) and the chart will update automatically.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
