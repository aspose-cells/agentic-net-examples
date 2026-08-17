// Title: Make a Shared Content Type Property Nillable in Aspose.Cells (C#)
// Description: Demonstrates how to mark a workbook as shared, add a custom content‑type property, set its IsNillable flag to true so the metadata becomes optional, and save the file as an .xlsx document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells IsNillable | shared content type property | optional workbook metadata | C# Excel custom property | set IsNillable true
// Common Searches: Aspose.Cells make content type property optional | C# set IsNillable for shared property | how to enable nillable metadata in Excel with Aspose | add shared property to workbook Aspose.Cells
// Developer Intent: Enable a shared custom property to be optional (nillable) across all generated workbooks.
// Use Cases: Create Excel files where a shared custom property can be omitted without breaking XML schema validation. | Generate multiple workbooks that reuse the same property but allow it to remain empty when not needed. | Export data with optional custom metadata for downstream processing or compliance.
// AI Prompts: Show a C# example that sets IsNillable = true for a shared content type property in Aspose.Cells. | Explain how to make a custom workbook property optional using the IsNillable flag. | Provide steps to apply the IsNillable setting to existing Aspose.Cells workbooks without recreating them.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Demonstrates how to mark a workbook as shared, add a custom content‑type property, set its IsNillable flag to true so the metadata becomes optional, and save the file as an .xlsx document using Aspose.Cells for .NET.
class SetIsNillableSharedProperty
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Mark the workbook as shared
        workbook.Settings.Shared = true;

        // Add a content type property named "Shared"
        int index = workbook.ContentTypeProperties.Add("Shared", "Aspose", "text");

        // Set IsNillable to true for the shared property
        workbook.ContentTypeProperties[index].IsNillable = true;

        // Save the workbook
        workbook.Save("SharedPropertyIsNillable.xlsx");
    }
}
