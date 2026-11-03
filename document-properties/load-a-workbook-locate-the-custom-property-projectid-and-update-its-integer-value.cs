// Title: How to update or add an integer custom document property (ProjectId) in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set integer custom document property in existing Excel file | Add or modify custom property named ProjectId with Aspose.Cells .NET | Update custom document property value in workbook and save using Aspose.Cells C# | Check if custom property exists and change its value in Excel using Aspose.Cells | Aspose.Cells create new integer custom property when missing in workbook
// Tags: Aspose.Cells custom document property manipulation | C# change numeric custom property in Excel workbook | Add new custom property when absent using Aspose.Cells | Workbook.CustomDocumentProperties example code | Persist workbook after property update with Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing Excel file into a Workbook, accesses its CustomDocumentProperties collection, checks for a property named "ProjectId", updates its integer value to 12345 if it exists or adds it if missing, and then saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Access the custom document properties collection
            var customProps = workbook.CustomDocumentProperties;

            // Check if the property "ProjectId" already exists
            var existingProp = customProps["ProjectId"];
            if (existingProp != null)
            {
                // Update the existing property's integer value
                existingProp.Value = 12345;
            }
            else
            {
                // Add a new custom property with the integer value
                customProps.Add("ProjectId", 12345);
            }

            // Save the workbook with the updated property
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
