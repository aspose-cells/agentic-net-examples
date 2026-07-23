// Title: C# – Update or Add the 'ProjectId' Custom Document Property in an Excel Workbook with Aspose.Cells
// Description: Load an existing .xlsx file, access its CustomDocumentPropertyCollection, set the integer value of the "ProjectId" property (or create it if missing), and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | custom document property | ProjectId | update Excel property | add Excel property | integer property | Workbook.Save | CustomDocumentPropertyCollection
// Common Searches: how to change a custom document property in Excel with Aspose.Cells C# | add missing custom property to workbook using Aspose.Cells | update integer property ProjectId in .xlsx via Aspose.Cells | retrieve and modify custom document properties C# Aspose | Aspose.Cells set custom property value programmatically
// Developer Intent: Modify the "ProjectId" custom document property in an Excel workbook, creating it when it does not already exist.
// Use Cases: Synchronize project IDs across a set of report workbooks by updating the ProjectId property. | Enforce a template rule that every generated workbook contains a ProjectId for downstream processing. | Automate version control by incrementing a numeric custom property in each exported Excel file.
// AI Prompts: Write C# code with Aspose.Cells that reads the "ProjectId" custom property from an existing workbook, sets it to a new integer, and adds the property if it is absent. | Show how to safely handle the exception when a custom document property does not exist and then create it using Aspose.Cells. | Explain a batch‑processing script that iterates over a folder of .xlsx files and updates each file's "ProjectId" property to a specified number with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Load an existing .xlsx file, access its CustomDocumentPropertyCollection, set the integer value of the "ProjectId" property (or create it if missing), and save the workbook using Aspose.Cells for .NET.
class UpdateProjectId
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of custom document properties
        CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

        // Try to retrieve the "ProjectId" property
        DocumentProperty projectProp = null;
        try
        {
            projectProp = customProps["ProjectId"];
        }
        catch
        {
            // Property does not exist
        }

        if (projectProp != null)
        {
            // Update the existing property with a new integer value
            projectProp.Value = 12345;
        }
        else
        {
            // If the property is missing, add it with the desired integer value
            customProps.Add("ProjectId", 12345);
        }

        // Save the workbook with the updated property
        workbook.Save("output.xlsx");
    }
}
