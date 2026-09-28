// Title: How to set a custom #DIV/0! error message in an Aspose.Cells PivotTable using C#
// AI Prompts: Generate C# code that creates a workbook, adds a PivotTable, enables DisplayErrorString, and assigns a custom ErrorString for division‑by‑zero errors. | Show the steps to calculate formulas, refresh the pivot, and save the workbook after customizing the error text in an Aspose.Cells PivotTable. | Provide a minimal example that demonstrates overriding the default #DIV/0! display with a user‑defined message in a PivotTable using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set custom error text for PivotTable division by zero | display custom #DIV/0! message in Excel pivot table with Aspose.Cells | how to use DisplayErrorString and ErrorString properties in Aspose.Cells PivotTable | replace default error string in Aspose.Cells pivot data field C# | sample code for customizing pivot table error messages in .NET
// Tags: Aspose.Cells PivotTable custom error string | DisplayErrorString property C# | ErrorString property Aspose.Cells | override #DIV/0! message Excel pivot | C# Aspose.Cells pivot error handling

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotErrorStringDemo
{
    // The example creates a workbook, fills cells with formulas that generate #DIV/0! errors, calculates the formulas, adds a PivotTable on range A1:C4, enables DisplayErrorString, sets a custom ErrorString ("Custom Division Error"), refreshes and calculates the pivot data, and saves the file as PivotTableErrorStringDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A: Category, Column B: Value, Column C: Formula that causes #DIV/0! error
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["C1"].PutValue("ErrorCalc");

            sheet.Cells["A2"].PutValue("X");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].Formula = "=B2/0"; // #DIV/0! error

            sheet.Cells["A3"].PutValue("Y");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["C3"].Formula = "=B3/0"; // #DIV/0! error

            sheet.Cells["A4"].PutValue("X");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C4"].Formula = "=B4/0"; // #DIV/0! error

            // Calculate formulas so that error values are materialized
            workbook.CalculateFormula();

            // Add a pivot table based on the data range A1:C4
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "ErrorPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure pivot fields: Category as row, ErrorCalc as data field
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Column A (Category)
            pivot.AddFieldToArea(PivotFieldType.Data, 2);  // Column C (ErrorCalc)

            // Enable custom error string display and set the custom message
            pivot.DisplayErrorString = true;
            pivot.ErrorString = "Custom Division Error";

            // Refresh data and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableErrorStringDemo.xlsx");

            Console.WriteLine("PivotTable created with custom error string.");
        }
    }
}
