// Title: How to add an integer custom document property named ProjectId to an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Open an existing Excel file or instantiate a new workbook, then insert a custom document property called ProjectId with an integer value via Aspose.Cells and save the result. | Modify the example to retrieve the ProjectId value from a configuration source or environment variable before adding it as a custom document property. | Enhance the code to detect whether a custom property named ProjectId already exists in the workbook and update its value instead of creating a duplicate.
// Common Searches: asp.net add integer custom document property to Excel using Aspose.Cells | c# aspose.cells set custom property ProjectId if workbook already contains it | how to open or create Excel file and add custom document property with Aspose.Cells .NET | aspose.cells example for integer custom document properties in Excel | save Excel workbook with custom property ProjectId using C#
// Tags: integer custom document property Aspose.Cells | create or load workbook Aspose.Cells C# | custom document properties Excel Aspose.Cells | save workbook with custom properties Aspose.Cells | check existing custom property before adding Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an existing Excel file or creates a new workbook, adds a custom document property named "ProjectId" with the integer value 12345, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook
            }

            // Add a custom document property named "ProjectId" with an integer value
            // Overload without isLinkToContent is used for compatibility with various Aspose.Cells versions
            workbook.CustomDocumentProperties.Add("ProjectId", 12345);

            // Save the workbook with the new custom property
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
