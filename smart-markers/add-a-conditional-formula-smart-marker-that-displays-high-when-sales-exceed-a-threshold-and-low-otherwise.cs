// Title: C# – Conditional IF Smart Marker in Aspose.Cells to Flag Sales as High or Low
// Description: This Aspose.Cells for .NET example creates a workbook, adds "Sales" and "Status" headers, and inserts a smart marker that uses the IF formula to output "High" when a sales value exceeds a configurable threshold (default 1000) and "Low" otherwise. A List<SalesData> is bound as the data source, processed with WorkbookDesigner, and the result is saved as an Excel file.
// Keywords: Aspose.Cells | C# | smart markers | conditional formula | IF function | sales threshold | WorkbookDesigner | Excel report automation | .NET | US developers | Europe developers
// Common Searches: Aspose.Cells conditional smart marker example | IF formula in smart markers C# | How to label sales high low using Aspose.Cells | Set threshold in Aspose.Cells smart marker | Generate status column with smart markers
// Developer Intent: Add a smart marker that evaluates an IF expression to display "High" or "Low" based on each Sales value.
// Use Cases: Automated sales reports that categorize each transaction as High or Low without manual formulas. | Excel dashboards that dynamically flag performance thresholds during data export. | Bulk data export pipelines where conditional text labels are required for downstream analytics.
// AI Prompts: Write C# code that inserts a conditional IF smart marker in Aspose.Cells to mark sales as High or Low with a customizable threshold. | Explain how to bind a collection of objects to a smart marker and process conditional formulas using WorkbookDesigner. | Show how to change the threshold value in the smart marker formula at runtime without modifying the source code.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace ConditionalFormulaSmartMarkerDemo
{
    // Simple data class for the smart marker data source
    // This Aspose.Cells for .NET example creates a workbook, adds "Sales" and "Status" headers, and inserts a smart marker that uses the IF formula to output "High" when a sales value exceeds a configurable threshold (default 1000) and "Low" otherwise. A List<SalesData> is bound as the data source, processed with WorkbookDesigner, and the result is saved as an Excel file.
    public class SalesData
    {
        public double Sales { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Add column headers
            sheet.Cells["A1"].PutValue("Sales");
            sheet.Cells["B1"].PutValue("Status");

            // 3. Insert a smart marker that evaluates a conditional formula.
            //    The marker will display "High" when the Sales value exceeds 1000,
            //    otherwise it will display "Low".
            //    Syntax: &="=IF(Sales>1000,\"High\",\"Low\")"
            sheet.Cells["B2"].PutValue("&=IF(Sales>1000,\"High\",\"Low\")");

            // 4. Prepare sample data source
            List<SalesData> data = new List<SalesData>
            {
                new SalesData { Sales = 750 },
                new SalesData { Sales = 1250 },
                new SalesData { Sales = 500 },
                new SalesData { Sales = 2000 }
            };

            // 5. Set up the WorkbookDesigner, assign the data source and process smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", data);
            designer.Process();

            // 6. Save the resulting workbook
            workbook.Save("ConditionalFormulaSmartMarkerOutput.xlsx");
        }
    }
}
