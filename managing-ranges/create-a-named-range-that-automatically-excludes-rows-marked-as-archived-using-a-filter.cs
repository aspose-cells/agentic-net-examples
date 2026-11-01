// Title: Define a dynamic named range that filters out rows with Status = "Archived" using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# create named range that excludes rows where Status equals Archived | How to apply the Excel FILTER formula in Aspose.Cells to build a dynamic named range | C# example for defining a named range based on a column condition with Aspose.Cells | Generate a named range that only includes active records in a workbook using Aspose.Cells | Use Excel 365 FILTER function in Aspose.Cells to hide archived rows
// Tags: dynamic named range using FILTER Aspose.Cells | conditional named range based on Status column C# | Aspose.Cells example for Excel 365 FILTER function | C# create named range that omits archived records | Excel workbook named range with column criteria Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a workbook, adds sample data with an ID, Status, Name, and Amount columns, and then defines a dynamic named range called "ActiveData". The range uses the Excel FILTER function to return only rows where the Status column is not "Archived". The named range is added to the workbook's name collection and saved as "NamedRange_ExcludeArchived.xlsx".
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Populate sample data (Header + rows)
        // Columns: ID, Status, Name, Amount
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Status");
        sheet.Cells["C1"].PutValue("Name");
        sheet.Cells["D1"].PutValue("Amount");

        // Sample rows
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Active");
        sheet.Cells["C2"].PutValue("Alice");
        sheet.Cells["D2"].PutValue(100);

        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Archived");
        sheet.Cells["C3"].PutValue("Bob");
        sheet.Cells["D3"].PutValue(200);

        sheet.Cells["A4"].PutValue(3);
        sheet.Cells["B4"].PutValue("Active");
        sheet.Cells["C4"].PutValue("Charlie");
        sheet.Cells["D4"].PutValue(150);

        // Define a dynamic named range "ActiveData" that excludes rows where Status = "Archived"
        // Using the FILTER function (available in Excel 365+). The range A2:D4 is the data area.
        // The condition B2:B4<>"Archived" filters out archived rows.
        string formula = "=FILTER(Data!$A$2:$D$4, Data!$B$2:$B$4<>\"Archived\")";

        // Add the named range to the workbook
        int nameIndex = workbook.Worksheets.Names.Add("ActiveData");
        Name activeName = workbook.Worksheets.Names[nameIndex];
        activeName.RefersTo = formula;

        // Save the workbook
        workbook.Save("NamedRange_ExcludeArchived.xlsx");
    }
}
