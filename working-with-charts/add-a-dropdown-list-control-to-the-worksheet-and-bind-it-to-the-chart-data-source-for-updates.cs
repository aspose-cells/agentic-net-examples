using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the list of options (dropdown items) in column A
        sheet.Cells["A2"].PutValue("Option1");
        sheet.Cells["A3"].PutValue("Option2");
        sheet.Cells["A4"].PutValue("Option3");
        sheet.Cells["A5"].PutValue("Option4");

        // Corresponding numeric values for the chart in column B
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["B5"].PutValue(40);

        // Cell that will hold the selected option from the dropdown
        sheet.Cells["C1"].PutValue("Option1"); // default selection

        // Add a ComboBox ActiveX control to the worksheet
        // Parameters: ControlType, upperLeftRow, upperLeftColumn, top, left, height, width
        Shape comboShape = sheet.Shapes.AddActiveXControl(
            ControlType.ComboBox, // type of control
            1,    // upper left row
            0,    // upper left column
            0,    // top offset (pixels)
            0,    // left offset (pixels)
            30,   // height (pixels)
            100); // width (pixels)

        // Cast to the specific ComboBox control to set its properties
        ComboBoxActiveXControl comboBox = (ComboBoxActiveXControl)comboShape.ActiveXControl;

        // Fill the ComboBox with the options defined in column A
        comboBox.ListFillRange = "A2:A5";

        // Link the selected value to cell C1; when the user picks an item,
        // C1 will be updated with the chosen option text
        comboBox.LinkedCell = "C1";

        // Add a column chart that will reflect the value associated with the selected option
        int chartIndex = sheet.Charts.Add(ChartType.Column, 10, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Use an INDEX/MATCH formula so the series data changes based on the dropdown selection
        // The formula returns the value from B2:B5 that matches the option in C1
        chart.NSeries.Add("=INDEX($B$2:$B$5, MATCH($C$1,$A$2:$A$5,0))", true);

        // Optional: set category labels (the list of all possible options)
        chart.NSeries.CategoryData = "A2:A5";

        // Save the workbook with the dropdown and linked chart
        workbook.Save("DropdownChartBinding.xlsx");
    }
}