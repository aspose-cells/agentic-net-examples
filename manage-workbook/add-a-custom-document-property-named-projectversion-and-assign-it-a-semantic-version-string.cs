// Title: Add a ProjectVersion custom document property (semantic version) to an Excel workbook using Aspose.Cells for C#
// Description: Creates a new Workbook, adds a custom document property named ProjectVersion with a semantic version string (e.g., 1.2.3), prints the value for verification, and saves the file as an XLSX workbook.
// Keywords: Aspose.Cells | custom document property | ProjectVersion | semantic version | C# | .NET | Excel workbook metadata | add property | save workbook | retrieve property
// Common Searches: Aspose.Cells add custom document property C# | How to set ProjectVersion property in Excel with Aspose | Read custom document property from Excel using Aspose.Cells | Save workbook after adding custom properties .NET | Add semantic version to Excel file programmatically
// Developer Intent: Add a ProjectVersion custom document property with a semantic version to a workbook and persist the change.
// Use Cases: Tag generated reports with the application version for auditability. | Embed build or release numbers into Excel templates used in CI/CD pipelines. | Store API version information in spreadsheets so downstream processes can verify compatibility. | Provide version metadata for automated document management systems.
// AI Prompts: Generate C# code with Aspose.Cells that adds a ProjectVersion custom property set to "2.0.0" and saves the workbook. | Show how to read the ProjectVersion custom property from an existing Excel file using Aspose.Cells. | Give best‑practice error‑handling patterns when adding or retrieving custom document properties in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, adds a custom document property named ProjectVersion with a semantic version string (e.g., 1.2.3), prints the value for verification, and saves the file as an XLSX workbook.
    public class AddCustomDocumentPropertyDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a custom document property named "ProjectVersion" with a semantic version string
                workbook.CustomDocumentProperties.Add("ProjectVersion", "1.2.3");

                // Display the added property to verify
                Console.WriteLine("ProjectVersion: " + workbook.CustomDocumentProperties["ProjectVersion"].Value);

                // Save the workbook to a file
                workbook.Save("ProjectVersionDemo.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AddCustomDocumentPropertyDemo.Run();
        }
    }
}
