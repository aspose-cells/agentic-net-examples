// Title: Set the DocumentVersion custom property to 2.0 in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom document property DocumentVersion to 2.0 | How to add or update a custom property in an existing Excel file with Aspose.Cells | C# code to create DocumentVersion property when it is missing in a workbook | Saving an Excel workbook after changing custom document properties using Aspose.Cells | Error handling for custom document properties in Aspose.Cells .NET
// Tags: Aspose.Cells metadata update | C# Excel workbook property modification | DocumentVersion handling with Aspose | modify workbook metadata .NET | save workbook after property update

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file (or creates a new one), ensures a custom document property named "DocumentVersion" exists, sets its value to "2.0", and saves the workbook, including error handling for property access and overall execution.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if the file exists; otherwise create a new workbook.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Ensure the custom property "DocumentVersion" exists, then set its value.
            try
            {
                var customProps = workbook.CustomDocumentProperties;
                var docVersionProp = customProps["DocumentVersion"];
                if (docVersionProp == null)
                {
                    customProps.Add("DocumentVersion", "2.0");
                }
                else
                {
                    docVersionProp.Value = "2.0";
                }
            }
            catch (Exception propEx)
            {
                Console.WriteLine($"Property error: {propEx.Message}");
            }

            // Save the workbook with the updated property.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
