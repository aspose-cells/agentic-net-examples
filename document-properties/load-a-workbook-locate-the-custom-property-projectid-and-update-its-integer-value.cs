// Title: Update or Add Custom Document Property "ProjectId" in Excel with Aspose.Cells for .NET
// Description: This example loads an Excel workbook, accesses its custom document properties, sets the integer value of ProjectId (creating the property if it does not exist), and saves the workbook to a new file.
// Keywords: Aspose.Cells | C# Excel | custom document property | ProjectId | update property | add property | .NET | Excel workbook metadata | set integer property | Aspose.Cells example
// Common Searches: Aspose.Cells update custom property C# | Add integer custom property to Excel workbook using Aspose.Cells | Check if custom document property exists Aspose.Cells .NET | Modify ProjectId property in Excel file with C# | How to create missing custom property in Aspose.Cells
// Developer Intent: Load an Excel workbook, locate or create the custom property named ProjectId, assign it an integer value, and save the changes.
// Use Cases: Batch‑update ProjectId across multiple workbooks for centralized reporting. | Embed a numeric version or identifier into Excel templates before distribution. | Ensure legacy workbooks comply with standards by adding a missing ProjectId property.
// AI Prompts: Write C# code with Aspose.Cells that reads the 'ProjectId' custom property, increments its integer value, and saves the workbook. | Show how to safely check for a 'ProjectId' custom property, add it with a default integer if absent, then update its value using Aspose.Cells. | Explain how to enumerate all custom document properties in a workbook and modify the one named 'ProjectId' with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// This example loads an Excel workbook, accesses its custom document properties, sets the integer value of ProjectId (creating the property if it does not exist), and saves the workbook to a new file.
class UpdateProjectIdProperty
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Try to get the custom document property named "ProjectId"
        DocumentProperty projectIdProp = workbook.CustomDocumentProperties["ProjectId"];

        if (projectIdProp != null)
        {
            // Update its value to a new integer (e.g., 2023)
            projectIdProp.Value = 2023;
        }
        else
        {
            // If the property does not exist, add it with the desired integer value
            workbook.CustomDocumentProperties.Add("ProjectId", 2023);
        }

        // Save the workbook with the updated property
        workbook.Save("output.xlsx");
    }
}
