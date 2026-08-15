// Title: Export Pivot Table Field Display Names to a Text File with Aspose.Cells for .NET
// Description: Shows how to generate a workbook, populate it with sample data, create a pivot table, refresh its data, and loop through the BaseFields collection to write each field's DisplayName into a plain‑text file before saving the workbook.
// Keywords: Aspose.Cells | .NET | C# | pivot table | BaseFields | DisplayName | export field names | write to text file | enumerate pivot fields | log pivot structure | Aspose.Cells API
// Common Searches: Aspose.Cells get pivot field names | C# write pivot table fields to file | list pivot table columns Aspose.Cells | enumerate BaseFields Aspose.Cells | export pivot field display names
// Developer Intent: Extract every pivot field's display identifier and persist it to a log file.
// Use Cases: Create an audit trail of a pivot table's schema for documentation. | Validate pivot configuration by comparing logged field identifiers with expected values. | Provide a simple field‑list report for business analysts without requiring Excel access.
// AI Prompts: Generate C# code that appends pivot field names to an existing log instead of overwriting it. | Show how to filter BaseFields by row, column, or data type before writing their names to a file. | Explain performance‑friendly techniques for logging field names from very large pivot tables using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotFieldLogger
{
    // Shows how to generate a workbook, populate it with sample data, create a pivot table, refresh its data, and loop through the BaseFields collection to write each field's DisplayName into a plain‑text file before saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Electronics";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Electronics";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 800;

            sheet.Cells["A4"].Value = "Furniture";
            sheet.Cells["B4"].Value = "North";
            sheet.Cells["C4"].Value = 600;

            sheet.Cells["A5"].Value = "Furniture";
            sheet.Cells["B5"].Value = "South";
            sheet.Cells["C5"].Value = 400;

            // Add a pivot table based on the sample data
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table so that fields are initialized
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Prepare the output file (overwrites if it already exists)
            string outputPath = "PivotFieldDisplayNames.txt";
            using (StreamWriter writer = new StreamWriter(outputPath, false))
            {
                // Iterate through all base fields of the pivot table
                foreach (PivotField field in pivotTable.BaseFields)
                {
                    // Log the display name of each field
                    writer.WriteLine(field.DisplayName);
                }
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("PivotFieldLoggerDemo.xlsx");
        }
    }
}
