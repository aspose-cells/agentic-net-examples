// Title: Set a Custom Calculation Error Message for an Aspose.Cells PivotTable (C#)
// Description: Demonstrates how to enable the DisplayErrorString flag and assign a custom ErrorString to an Aspose.Cells PivotTable in C#. The example creates a workbook, adds sample data, builds a pivot table, configures the custom error message, recalculates, and saves the file, allowing you to replace default Excel error codes with a user‑defined text.
// Keywords: Aspose.Cells PivotTable custom error message | DisplayErrorString property C# | ErrorString Aspose.Cells | pivot table calculation error handling | replace #DIV/0! with custom text | Aspose.Cells .NET pivot table example
// Common Searches: Aspose.Cells set custom error string for PivotTable | DisplayErrorString example C# | How to change pivot table error text in Aspose.Cells | ErrorString property usage Aspose.Cells | C# pivot table custom error handling Aspose
// Developer Intent: Configure a PivotTable to show a user‑defined message instead of default calculation errors.
// Use Cases: Show a friendly message (e.g., "Data unavailable") instead of #DIV/0! in sales dashboards. | Standardize error text across multiple automatically generated reports. | Highlight data quality issues in Excel files produced by a .NET reporting service.
// AI Prompts: Write C# code that creates a PivotTable with Aspose.Cells and sets DisplayErrorString to true with a custom ErrorString. | Explain how DisplayErrorString and ErrorString affect pivot table error display in Aspose.Cells. | List the verification steps to confirm the custom error message appears after recalculating the PivotTable.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to enable the DisplayErrorString flag and assign a custom ErrorString to an Aspose.Cells PivotTable in C#. The example creates a workbook, adds sample data, builds a pivot table, configures the custom error message, recalculates, and saves the file, allowing you to replace default Excel error codes with a user‑defined text.
    public class PivotTableDisplayErrorStringDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Sales";
            cells["A2"].Value = "Bike";
            cells["B2"].Value = 1000;
            cells["A3"].Value = "Car";
            cells["B3"].Value = 2000;
            cells["A4"].Value = "Bike";
            cells["B4"].Value = 1500;
            cells["A5"].Value = "Car";
            cells["B5"].Value = 2500;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table (optional: add fields)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Sales as data field

            // Enable custom error string display and set the custom message
            pivotTable.DisplayErrorString = true;
            pivotTable.ErrorString = "Custom Calculation Error";

            // Recalculate the pivot table to apply the settings
            pivotTable.CalculateData();

            // Output the current settings to the console (for verification)
            Console.WriteLine("DisplayErrorString: " + pivotTable.DisplayErrorString);
            Console.WriteLine("ErrorString: " + pivotTable.ErrorString);

            // Save the workbook to a file
            string outputPath = "PivotTableDisplayErrorStringDemo_out.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
    }
}
