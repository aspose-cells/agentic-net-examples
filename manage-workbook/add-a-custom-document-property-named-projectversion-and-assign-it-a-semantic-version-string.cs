// Title: Add a custom document property named ProjectVersion with a semantic version string to an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new Workbook, adds a user-defined property with a semantic version string (e.g., "1.2.3"), and saves it as Output.xlsx using Aspose.Cells. | Write C# using Aspose.Cells to open an existing Excel file, insert or update a user-defined property containing a semantic version value, and persist the changes.
// Common Searches: aspnet how to set a custom document property called ProjectVersion in an Excel file with Aspose.Cells | c# add semantic version string as custom property to workbook using Aspose.Cells | using Aspose.Cells to store project version in Excel workbook metadata | example code for adding custom document properties to a new workbook in C# Aspose.Cells
// Tags: add custom document property Aspose.Cells C# | store version string in Excel workbook metadata | Aspose.Cells workbook custom properties manipulation | C# set custom property in Excel file | Aspose.Cells add custom property to workbook

using Aspose.Cells;
using System;

// Demonstrates creating a new Workbook, adding a custom document property named ProjectVersion with the semantic version "1.2.3", and saving the file as Output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a custom document property named "ProjectVersion" with a semantic version string
        workbook.CustomDocumentProperties.Add("ProjectVersion", "1.2.3");

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
