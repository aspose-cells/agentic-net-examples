// Title: Mark a custom DocumentProperty (ProjectId) as optional by setting IsNillable = true with Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a "ProjectId" custom document property to an Excel workbook and attempts to set its IsNillable flag to true using Aspose.Cells. | Explain why Aspose.Cells DocumentProperty does not expose an IsNillable setter and suggest alternative patterns for representing optional custom properties in .NET.
// Common Searches: asp.net aspose.cells make custom document property optional | c# aspose.cells IsNillable flag not available for custom properties | how to create nullable custom Excel property with aspose.cells | aspose.cells add custom property and set it as nillable | set optional custom document property in excel using aspose.cells c#
// Tags: Aspose.Cells custom document property optional flag | C# set IsNillable Aspose.Cells | nullable custom Excel property Aspose.Cells | DocumentProperty nillable limitation Aspose | Excel custom property optional handling .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads or creates an Excel workbook, accesses its CustomDocumentProperties collection, adds a "ProjectId" property if missing, and notes that Aspose.Cells' DocumentProperty class does not expose an IsNillable flag, so true optionality cannot be set directly before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input workbook exists; create a new one if it doesn't.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook
            }

            // Access custom document properties.
            var customProps = workbook.CustomDocumentProperties;

            // Retrieve or add the "ProjectId" property.
            var projectIdProp = customProps.Contains("ProjectId")
                ? customProps["ProjectId"]
                : customProps.Add("ProjectId", string.Empty);

            // Note: Aspose.Cells DocumentProperty does not expose an IsNillable flag.
            // If needed, you can set the value or leave it as is.

            // Save the workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
