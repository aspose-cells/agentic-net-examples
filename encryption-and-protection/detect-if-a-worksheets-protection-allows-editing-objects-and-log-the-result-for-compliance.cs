// Title: Detect Worksheet AllowEditingObject Flag with Aspose.Cells for .NET
// Description: Loads an Excel file, reads the worksheet's Protection.AllowEditingObject property, logs whether drawing objects can be edited, and saves the workbook unchanged—useful for compliance checks.
// Keywords: Aspose.Cells worksheet protection | AllowEditingObject C# | Excel object editing permission | read worksheet protection flag | .NET Excel security audit | Aspose.Cells compliance example
// Common Searches: Aspose.Cells read AllowEditingObject | check if Excel sheet allows editing objects | C# worksheet protection flag Aspose | log worksheet protection settings .NET | Excel security audit with Aspose.Cells
// Developer Intent: Identify whether a worksheet’s protection permits editing of drawing objects and output the result.
// Use Cases: Verify that published workbooks block object editing for regulatory compliance. | Generate a quick report of the AllowEditingObject status across all sheets in a workbook. | Skip further processing on sheets that allow object modifications.
// AI Prompts: Create a C# routine that iterates through every worksheet in a workbook and records the AllowEditingObject value using Aspose.Cells. | Show how to disable object editing on a protected worksheet and save the changes with Aspose.Cells. | Explain how to combine AllowEditingObject with other protection flags to perform a full Excel security audit.

using System;
using Aspose.Cells;

// Loads an Excel file, reads the worksheet's Protection.AllowEditingObject property, logs whether drawing objects can be edited, and saves the workbook unchanged—useful for compliance checks.
class WorksheetProtectionCheck
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (or specify by index/name)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the protection settings for the worksheet
        Protection protection = worksheet.Protection;

        // Check whether editing of drawing objects is allowed
        bool allowEditingObject = protection.AllowEditingObject;

        // Log the compliance result
        Console.WriteLine($"Worksheet \"{worksheet.Name}\" AllowEditingObject: {allowEditingObject}");

        // Save the workbook if needed (no modifications made in this example)
        workbook.Save("output.xlsx");
    }
}
