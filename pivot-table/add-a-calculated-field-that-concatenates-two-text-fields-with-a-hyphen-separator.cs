// Title: Create a hyphen‑separated FullName calculated field in an Aspose.Cells pivot table using C#
// AI Prompts: Generate a C# program that builds a workbook, adds FirstName and LastName columns, creates a pivot table, and defines a calculated field named FullName using the formula =FirstName & "-" & LastName. | Show how to refresh the pivot cache and recalculate the pivot table after inserting a calculated field with Aspose.Cells for .NET. | Explain how to save the workbook as a .xlsx file after adding the concatenated FullName field to the pivot table.
// Common Searches: Aspose.Cells C# add calculated field that joins two text columns with a hyphen in a pivot table | How to use AddCalculatedField to concatenate FirstName and LastName in an Aspose.Cells pivot table | Refresh pivot data after adding a calculated field in Aspose.Cells .NET | Save workbook with custom calculated field in Aspose.Cells pivot table example
// Tags: AddCalculatedField concatenation Aspose.Cells | pivot table string concatenation formula C# | refresh pivot cache Aspose.Cells | save workbook with calculated field .xlsx | Aspose.Cells calculated field hyphen separator

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, populates FirstName and LastName columns, adds a pivot table at D3, places the name fields in the row area, defines a calculated field called FullName that concatenates the two columns with a hyphen using the '&' operator, refreshes and recalculates the pivot cache, and finally saves the file as PivotTable_With_ConcatenatedField.xlsx.
    public class AddCalculatedFieldConcatenateDemo
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

            // Populate sample data (FirstName and LastName)
            sheet.Cells["A1"].PutValue("FirstName");
            sheet.Cells["B1"].PutValue("LastName");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue("Doe");
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue("Smith");
            sheet.Cells["A4"].PutValue("Alice");
            sheet.Cells["B4"].PutValue("Brown");

            // Add a pivot table based on the data range, placed at D3
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add FirstName and LastName fields to the Row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "FirstName");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "LastName");

            // Add a calculated field that concatenates FirstName and LastName with a hyphen
            // In pivot calculated fields, '&' is used for string concatenation
            string formula = "=FirstName & \"-\" & LastName";
            pivotTable.AddCalculatedField("FullName", formula, true); // dragToDataArea = true

            // Refresh the pivot cache and calculate the pivot data
            pivotTable.RefreshData();      // Correct API to refresh source data
            pivotTable.CalculateData();    // Recalculate the pivot table

            // Save the workbook
            try
            {
                workbook.Save("PivotTable_With_ConcatenatedField.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}
