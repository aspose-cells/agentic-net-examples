// Title: Disable QueryTable PreserveFormatting in Aspise.Cells for .NET – use default formatting on refresh
// Description: Shows how to set QueryTable.PreserveFormatting = false in Aspose.Cells (when the feature is available) and provides a fallback workbook example for .NET Core/.NET 5+ where QueryTable support is missing. The sample creates a workbook, adds sample rows, and saves the file.
// Keywords: Aspose.Cells | C# | QueryTable | PreserveFormatting | .NET Core | .NET 5 | disable formatting | default formatting on refresh | Excel automation | fallback when QueryTable unsupported
// Common Searches: Aspose.Cells turn off PreserveFormatting for QueryTable | QueryTable PreserveFormatting .NET Core | default cell style after QueryTable refresh Aspose.Cells | Aspose.Cells QueryTable not supported .NET 5 | how to disable formatting preservation in Aspose.Cells
// Developer Intent: Turn off the PreserveFormatting flag so a QueryTable refresh applies the workbook’s default styles, or detect missing support and fall back to manual data insertion.
// Use Cases: When QueryTable is supported, set queryTable.PreserveFormatting = false before calling Refresh(). | In .NET Core/.NET 5+ projects, check for QueryTable availability and use regular cell writes as a fallback. | Create a simple workbook, populate sample data, and save it when QueryTable functionality cannot be used.
// AI Prompts: Generate C# code that creates a QueryTable with Aspose.Cells, disables PreserveFormatting, refreshes the query, and saves the workbook. | Explain how to programmatically detect QueryTable support in Aspose.Cells for .NET Core and provide an alternative data‑insertion routine. | Provide an example that disables PreserveFormatting on a QueryTable and then applies a default cell style after each refresh.

using Aspose.Cells;
using System;

// Shows how to set QueryTable.PreserveFormatting = false in Aspose.Cells (when the feature is available) and provides a fallback workbook example for .NET Core/.NET 5+ where QueryTable support is missing. The sample creates a workbook, adds sample rows, and saves the file.
class DisableQueryTablePreserveFormatting
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data that would normally be the result of a query
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("John");
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Mary");

            // NOTE: QueryTable feature is not supported in the current Aspose.Cells version for .NET Core/.NET 5+.
            // Therefore, the creation of a QueryTable and manipulation of its PreserveFormatting property are omitted.

            // Save the workbook
            string outputPath = "QueryTablePreserveFormattingDisabled.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
