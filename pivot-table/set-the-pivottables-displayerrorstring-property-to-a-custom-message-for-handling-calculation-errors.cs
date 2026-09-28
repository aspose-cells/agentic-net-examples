// Title: Set a custom calculation error message for a PivotTable using Aspose.Cells in C#
// AI Prompts: Write C# code that creates a workbook, fills it with sample data, adds a PivotTable, enables DisplayErrorString, assigns a custom ErrorString, recalculates the pivot, and saves the file. | Show how to configure an Aspose.Cells PivotTable to display a specific error text when a calculation fails, using the DisplayErrorString and ErrorString properties.
// Common Searches: Aspose.Cells C# how to display a custom error string in a PivotTable | example of using DisplayErrorString property with Aspose.Cells PivotTable | set custom calculation error message for PivotTable in .NET | C# Aspose.Cells pivot table error handling custom message | show 'Calculation Error' in Aspose.Cells PivotTable output
// Tags: Aspose.Cells PivotTable custom error string | C# DisplayErrorString property Aspose.Cells | PivotTable ErrorString configuration .NET | Aspose.Cells calculation error handling in PivotTable | generate workbook with custom pivot error message C# | Aspose.Cells set custom error text for pivot data

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, adding sample data, inserting a PivotTable, enabling DisplayErrorString, setting a custom ErrorString, recalculating the pivot, and saving the workbook using Aspose.Cells for .NET.
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
                Console.WriteLine($"Error: {ex.Message}");
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

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Sales as data field

            // Enable custom error string display and set the custom message
            pivotTable.DisplayErrorString = true;
            pivotTable.ErrorString = "Calculation Error";

            // Force calculation of the pivot table to apply the settings
            pivotTable.CalculateData();

            // Save the workbook to a file
            string outputPath = "PivotTableDisplayErrorStringDemo_out.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
