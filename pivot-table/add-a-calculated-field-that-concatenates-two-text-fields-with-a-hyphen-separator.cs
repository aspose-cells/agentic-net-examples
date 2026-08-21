// Title: C# – Add a Hyphen‑Separated FullName Calculated Field to an Aspose.Cells Pivot Table
// Description: This example shows how to create a workbook with FirstName and LastName columns, build a pivot table on range A1:B4, and use PivotTable.AddCalculatedField to add a new field called FullName that concatenates the two text columns with a hyphen ("FirstName & \"-\" & LastName"). The pivot is refreshed, calculated, and saved as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | pivot table | calculated field | concatenate text columns | FullName field | hyphen separator | AddCalculatedField | Excel formula in Aspose | RefreshData | CalculateData | GitHub example | global
// Common Searches: Aspose.Cells add calculated field C# | concatenate first and last name in pivot table | FullName column in Aspose pivot | hyphen separator calculated field Aspose | how to use AddCalculatedField in .NET
// Developer Intent: Generate a pivot‑table calculated column that joins FirstName and LastName with a dash using Aspose.Cells for .NET.
// Use Cases: Create a read‑only FullName label for employee reports without altering source data. | Show combined names as row headers in a pivot view for clearer analysis. | Export pivot tables that include a concatenated name field for downstream BI tools.
// AI Prompts: Write C# code with Aspose.Cells to add a calculated field that merges two text columns using a hyphen in a pivot table. | Explain the syntax of AddCalculatedField for string concatenation in an Aspose.Cells pivot. | Demonstrate how to refresh and calculate a pivot table after inserting a concatenated calculated field in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCalculatedFieldDemo
{
    // This example shows how to create a workbook with FirstName and LastName columns, build a pivot table on range A1:B4, and use PivotTable.AddCalculatedField to add a new field called FullName that concatenates the two text columns with a hyphen ("FirstName & \"-\" & LastName"). The pivot is refreshed, calculated, and saved as an XLSX file using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with two text columns: FirstName and LastName
            sheet.Cells["A1"].PutValue("FirstName");
            sheet.Cells["B1"].PutValue("LastName");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue("Doe");
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue("Smith");
            sheet.Cells["A4"].PutValue("Alice");
            sheet.Cells["B4"].PutValue("Brown");

            // Create a pivot table based on the data range A1:B4
            // Destination top‑left cell for the pivot table is E3
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "EmployeePivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the text fields to the Row area so they appear in the pivot view
            pivotTable.AddFieldToArea(PivotFieldType.Row, "FirstName");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "LastName");

            // Add a calculated field that concatenates FirstName and LastName with a hyphen
            // Excel formula syntax: =FirstName & "-" & LastName
            // The second overload automatically drags the field to the Data area
            pivotTable.AddCalculatedField("FullName", "=FirstName & \"-\" & LastName");

            // Refresh and calculate the pivot table to ensure the calculated field is evaluated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTable_WithFullNameCalculatedField.xlsx");
        }
    }
}
