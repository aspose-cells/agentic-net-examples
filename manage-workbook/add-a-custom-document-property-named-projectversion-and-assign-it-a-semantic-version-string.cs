// Title: Add a custom document property named ProjectVersion with a semantic version string to an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to add a custom document property to an Excel file with Aspose.Cells in C# | set a semantic version value as a custom property in a workbook using Aspose.Cells | Aspose.Cells C# example for adding ProjectVersion property to a new workbook | saving an Excel workbook with custom metadata using Aspose.Cells .NET | create and save an .xlsx file with custom document properties in C#
// Tags: user-defined property addition Aspose.Cells | semantic version string Excel metadata | Aspose.Cells workbook metadata manipulation C# | generate .xlsx with extra metadata using .NET | manage workbook custom fields Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new Workbook, adds a custom document property called ProjectVersion with the semantic version "1.2.3", and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a custom document property named "ProjectVersion" with a semantic version string
        workbook.CustomDocumentProperties.Add("ProjectVersion", "1.2.3");

        // Save the workbook to a file
        workbook.Save("Output.xlsx", SaveFormat.Xlsx);
    }
}
