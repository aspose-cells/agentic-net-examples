// Title: C# – List All Pivot Table Fields and Their Display Names with Aspose.Cells
// Description: Creates a workbook, adds sample data, builds a PivotTable, refreshes it, then iterates through the PivotTable.BaseFields collection to print each field’s DisplayName (or Name when empty) to the console. The workbook is saved as a demonstration of field enumeration.
// Keywords: Aspose.Cells C# pivot table | enumerate PivotTable fields | BaseFields collection | display name of pivot fields | list pivot field names .NET | Aspose.Cells API example | iterate over pivot fields | C# get pivot field display name | Aspose.Cells PivotTable enumeration | print pivot field names
// Common Searches: How to loop through pivot fields in Aspose.Cells | Get display name of each pivot field C# | Aspose.Cells BaseFields example | Retrieve pivot table field names programmatically | Aspose.Cells enumerate pivot fields .NET
// Developer Intent: Retrieve and print every pivot field’s display name from a PivotTable.
// Use Cases: Verify required fields exist before configuring a pivot layout | Populate a UI dropdown with available pivot fields for user selection | Log field names for debugging pivot table setup | Generate documentation of pivot schema automatically | Create dynamic reports that adapt to changing source columns
// AI Prompts: Write C# code using Aspose.Cells to enumerate all PivotTable fields and output their DisplayName, falling back to Name when DisplayName is empty. | Explain how to access the BaseFields collection of a PivotTable and retrieve both Name and DisplayName properties for each field. | Show how to change the DisplayName of pivot fields after iterating through them in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // Creates a workbook, adds sample data, builds a PivotTable, refreshes it, then iterates through the PivotTable.BaseFields collection to print each field’s DisplayName (or Name when empty) to the console. The workbook is saved as a demonstration of field enumeration.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Product";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Electronics";
            sheet.Cells["B2"].Value = "Laptop";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Electronics";
            sheet.Cells["B3"].Value = "Phone";
            sheet.Cells["C3"].Value = 800;

            sheet.Cells["A4"].Value = "Furniture";
            sheet.Cells["B4"].Value = "Chair";
            sheet.Cells["C4"].Value = 150;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table (row, column, data)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table so that fields are initialized
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Iterate over all base pivot fields and output each field's display name
            PivotFieldCollection allFields = pivotTable.BaseFields;
            Console.WriteLine("Pivot Fields Display Names:");
            foreach (PivotField field in allFields)
            {
                // DisplayName may be empty if not set; fallback to Name
                string displayName = string.IsNullOrEmpty(field.DisplayName) ? field.Name : field.DisplayName;
                Console.WriteLine($"- {displayName}");
            }

            // Save the workbook (optional, just to complete lifecycle)
            workbook.Save("PivotFieldsDisplayNamesDemo.xlsx");
        }
    }
}
