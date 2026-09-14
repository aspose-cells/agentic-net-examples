// Title: Update or add the integer custom document property "ProjectId" in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, checks for a custom property named ProjectId, assigns a new integer value, and saves the workbook. | Write a .NET snippet that adds a ProjectId integer custom document property to an Excel file when it is missing, otherwise updates its value, using Aspose.Cells.
// Common Searches: asp.net aspose.cells change integer custom property in existing Excel file | c# update ProjectId custom document property in workbook | add missing custom property to Excel workbook programmatically Aspose.Cells | how to set custom document property value with Aspose.Cells .NET | modify custom properties of an .xlsx using Aspose.Cells C#
// Tags: update custom document property Aspose.Cells | add integer custom property C# Excel | ProjectId custom property Aspose.Cells | modify workbook custom properties .NET | set custom document property value Excel

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook (input.xlsx) with Aspose.Cells, accesses its CustomDocumentProperties collection, updates the integer value of the "ProjectId" property to 12345 (or adds it if missing), and saves the result as output.xlsx, handling missing files and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of custom document properties
            var customProps = workbook.CustomDocumentProperties;

            // Attempt to retrieve the "ProjectId" property
            var projectIdProp = customProps["ProjectId"];

            if (projectIdProp != null)
            {
                // Update the existing property's integer value
                projectIdProp.Value = 12345;
            }
            else
            {
                // Add the property if it does not exist
                customProps.Add("ProjectId", 12345);
            }

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
