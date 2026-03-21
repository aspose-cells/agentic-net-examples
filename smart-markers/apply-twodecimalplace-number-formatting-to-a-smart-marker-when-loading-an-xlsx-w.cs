using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerNumberFormatting
{
    public class Program
    {
        public static void Main()
        {
            // Load the XLSX workbook that contains smart markers
            Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

            // Assume the smart marker is placed in cell B2 of the first worksheet
            // Example smart marker: &=$Amount
            Worksheet sheet = workbook.Worksheets[0];
            Cell smartMarkerCell = sheet.Cells["B2"];

            // Create a style that formats numbers with two decimal places (0.00)
            Style twoDecimalStyle = workbook.CreateStyle();
            twoDecimalStyle.Number = 2; // 2 corresponds to the built‑in "0.00" format

            // Apply the style to the cell containing the smart marker
            smartMarkerCell.SetStyle(twoDecimalStyle);

            // Prepare a data source for the smart marker
            // Here we use a simple list of anonymous objects with an Amount property
            var data = new List<object>
            {
                new { Amount = 123.456 },
                new { Amount = 78.9 },
                new { Amount = 0.5 }
            };

            // Set up the WorkbookDesigner, assign the data source, and process the smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", data);
            designer.Process();

            // Save the resulting workbook
            workbook.Save("ResultWithTwoDecimalFormatting.xlsx");
        }
    }
}