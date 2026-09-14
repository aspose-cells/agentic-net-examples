// Title: Enable DisplayErrorString to show '#N/A' for missing data in an Aspose.Cells PivotTable (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a pivot table and configures it to display '#N/A' for null entries by turning on the error‑display option and assigning a custom error string. | Update an existing Aspose.Cells workbook in C# to activate the error‑value flag for a pivot table and set a custom placeholder for missing data.
// Common Searches: Aspose.Cells C# pivot table display '#N/A' for null values | How to enable error display for missing data in Aspose.Cells pivot tables | Set custom error placeholder for missing values in Aspose.Cells pivot table | Example of configuring error string in Aspose.Cells pivot table C# | Pivot table null data handling using Aspose.Cells
// Tags: Aspose.Cells pivot table DisplayErrorString | C# configure error string Aspose.Cells | show #N/A error values pivot Aspose.Cells | pivot table missing data handling Aspose.Cells | Aspose.Cells error display flag C#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace Demo
{
    // The example creates a workbook, adds sample data with a null entry, builds a pivot table, enables the DisplayErrorString option, sets the ErrorString to "#N/A", refreshes and calculates the pivot, and saves the file as PivotTableShowErrorValues.xlsx.
    class ShowErrorValuesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with a missing value (null) to generate an error in the pivot
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(null); // missing data point
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Configure the pivot fields (row and data)
                pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
                pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

                // Enable custom error string display and set it to "#N/A"
                pivot.DisplayErrorString = true;
                pivot.ErrorString = "#N/A";

                // Refresh source data and calculate the pivot table
                pivot.RefreshData();      // correct API call
                pivot.CalculateData();

                // Save the workbook with the configured pivot table
                workbook.Save("PivotTableShowErrorValues.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ShowErrorValuesDemo.Run();
        }
    }
}
