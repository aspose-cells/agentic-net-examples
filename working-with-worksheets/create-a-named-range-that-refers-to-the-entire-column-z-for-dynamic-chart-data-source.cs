// Title: Aspose.Cells .NET: Create a named range for the entire column Z to use as a dynamic chart source
// Description: This example shows how to create a new workbook, rename the first worksheet to "Data", add a named range called "ColumnZ", set its RefersTo formula to the whole column Z ("=Data!$Z:$Z"), and save the file. The named range can then be referenced in chart series for a dynamic data source that automatically expands with new rows.
// Keywords: Aspose.Cells | C# | .NET | named range entire column | column Z named range | RefersTo formula | dynamic chart data source | Excel automation | chart series | workbook save
// Common Searches: Aspose.Cells create named range for whole column | C# named range column Z Aspose.Cells | dynamic chart source using named range Aspose.Cells | set RefersTo to $Z:$Z in Aspose.Cells | how to bind column named range to chart series .NET
// Developer Intent: Add a named range that points to the full column Z so it can serve as a self‑updating data source for Excel charts created with Aspose.Cells.
// Use Cases: Define ColumnZ once and reuse it across multiple charts to ensure all new entries are included automatically. | Switch the named range to another worksheet without changing the column reference, keeping chart formulas intact. | Combine the column‑wide named range with conditional formatting or data validation that applies to the entire column.
// AI Prompts: Generate C# code using Aspose.Cells to create a named range that references the entire column Z and attach it to a chart series. | Show how to update an existing named range to point to a different worksheet while keeping the column reference unchanged. | Provide a complete Aspose.Cells example that creates a line chart using a column‑wide named range as its data source.

using System;
using Aspose.Cells;

// This example shows how to create a new workbook, rename the first worksheet to "Data", add a named range called "ColumnZ", set its RefersTo formula to the whole column Z ("=Data!$Z:$Z"), and save the file. The named range can then be referenced in chart series for a dynamic data source that automatically expands with new rows.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and give it a friendly name
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Add a named range that will refer to the entire column Z
        int nameIndex = workbook.Worksheets.Names.Add("ColumnZ");
        Name columnZ = workbook.Worksheets.Names[nameIndex];

        // Set the RefersTo formula to the whole column Z on the sheet named "Data"
        // The syntax "$Z:$Z" denotes the entire column Z
        columnZ.RefersTo = "=Data!$Z:$Z";

        // The named range can now be used as a dynamic data source for charts, e.g.:
        // chart.NSeries.Add("ColumnZ", true);

        // Save the workbook
        workbook.Save("ColumnZNamedRange.xlsx");
    }
}
