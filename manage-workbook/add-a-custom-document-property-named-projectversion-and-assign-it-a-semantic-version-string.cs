// Title: Add a custom document property named ProjectVersion with a semantic version to an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add custom document property with version number | How to store semantic version in an Excel file using Aspose.Cells | Set ProjectVersion custom property in .xlsx via .NET | Create workbook and add custom metadata property with Aspose.Cells | Add custom property to Excel workbook programmatically in C#
// Tags: Aspose.Cells workbook metadata addition | ProjectVersion Excel custom metadata | semantic version string .xlsx | C# set workbook custom attribute | add custom attribute using Aspose.Cells API

using Aspose.Cells;
using System;

// Demonstrates creating a new Workbook, adding a custom document property called ProjectVersion containing a semantic version string (e.g., "1.2.3"), and saving the workbook as Output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();

        // Add a custom document property named "ProjectVersion" with a semantic version string
        workbook.CustomDocumentProperties.Add("ProjectVersion", "1.2.3");

        // Save the workbook to a file (lifecycle save)
        workbook.Save("Output.xlsx");
    }
}
