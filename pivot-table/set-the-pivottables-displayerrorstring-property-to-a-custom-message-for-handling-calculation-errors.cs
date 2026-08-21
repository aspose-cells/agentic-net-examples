// Title: Aspose.Cells C# – Set a custom calculation error message for a PivotTable with DisplayErrorString
// Description: Creates a workbook, adds sample data, builds a PivotTable, enables DisplayErrorString, assigns a custom ErrorString, recalculates the table, and saves the file. Shows how to replace default Excel errors with a user‑defined message in C#.
// Keywords: Aspose.Cells PivotTable custom error | DisplayErrorString property | ErrorString Aspose.Cells | C# pivot table error handling | Aspose.Cells .NET example | Excel calculation error message | PivotTable DisplayErrorString C# | Aspose.Cells US developers
// Common Searches: Aspose.Cells set custom error text for PivotTable | DisplayErrorString example C# | How to change #DIV/0! message in Aspose.Cells | PivotTable ErrorString property usage | C# code to show custom error in Excel pivot
// Developer Intent: Apply a user‑defined error string to a PivotTable so calculation errors display a custom message.
// Use Cases: Provide end‑users with friendly messages instead of Excel error codes in generated reports. | Standardize error handling across multiple pivot tables in a workbook. | Hide sensitive calculation details by substituting them with a custom placeholder.
// AI Prompts: Write C# code that creates a PivotTable with Aspose.Cells and sets DisplayErrorString to true with a custom ErrorString. | Explain the impact of DisplayErrorString and ErrorString on PivotTable calculations and how to refresh the table after changes. | Give a step‑by‑step tutorial for adding a custom error message to an existing PivotTable in an Excel file using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, builds a PivotTable, enables DisplayErrorString, assigns a custom ErrorString, recalculates the table, and saves the file. Shows how to replace default Excel errors with a user‑defined message in C#.
    public class PivotTableDisplayErrorStringDemo
    {
        // Entry point required for console application
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
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field

            // Enable custom error string display and set the custom message
            pivotTable.DisplayErrorString = true;
            pivotTable.ErrorString = "Custom Calculation Error";

            // Recalculate the pivot table to apply the settings
            pivotTable.CalculateData();

            // Output the current settings to the console (for verification)
            Console.WriteLine("DisplayErrorString: " + pivotTable.DisplayErrorString);
            Console.WriteLine("ErrorString: " + pivotTable.ErrorString);

            // Save the workbook to a file
            workbook.Save("PivotTableDisplayErrorStringDemo_out.xlsx");
        }
    }
}
