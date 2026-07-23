// Title: C# – Verify and Add the "ClientName" Custom Document Property with Aspose.Cells
// Description: Loads an Excel workbook using Aspose.Cells, checks whether a custom document property called "ClientName" already exists, adds the property with a sample value if it is missing, and saves the updated file.
// Keywords: Aspose.Cells C# custom property | Workbook.CustomDocumentProperties.Contains | add custom document property Excel | check existing property Aspose | Excel metadata manipulation .NET
// Common Searches: Aspose.Cells check if custom property exists | add custom document property only when absent C# | prevent duplicate custom properties Excel Aspose | how to read and write workbook metadata with Aspose.Cells
// Developer Intent: Load an Excel file, determine if the "ClientName" custom property is present, and create it only when it does not already exist.
// Use Cases: Automatically embed client identifiers into report workbooks before distribution. | Batch‑process a library of spreadsheets, ensuring each file contains required metadata without creating duplicates. | Integrate metadata validation into a document‑management pipeline that enriches files with missing custom properties.
// AI Prompts: Generate C# code with Aspose.Cells that checks for a "ProjectId" custom property and adds a GUID value if the property is absent. | Create a reusable method that accepts a file path, property name, and value, adds the custom document property only when missing, and handles file‑not‑found exceptions. | Write a script that scans a directory of Excel files, verifies the presence of a "ReviewedBy" custom property, and inserts a default reviewer name where the property is missing.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Loads an Excel workbook using Aspose.Cells, checks whether a custom document property called "ClientName" already exists, adds the property with a sample value if it is missing, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Verify whether the custom document property "ClientName" already exists
        bool hasClientName = workbook.CustomDocumentProperties.Contains("ClientName");

        // If the property does not exist, add it with a sample value
        if (!hasClientName)
        {
            // Add a new custom property of type string
            workbook.CustomDocumentProperties.Add("ClientName", "Acme Corp");
        }

        // Save the workbook (overwrites or creates a new file)
        workbook.Save("output.xlsx");
    }
}
