// Title: Set a custom DocumentVersion property to "2.0" in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# method that takes a workbook path and a version string, opens the file with Aspose.Cells, adds or updates the custom document property "DocumentVersion" with the supplied version, and saves the workbook. | Refactor the example to read the version value from an appsettings.json file and apply it to the workbook's DocumentVersion custom property using Aspose.Cells.
// Common Searches: Aspose.Cells C# set custom document property DocumentVersion to specific value | How to add or update a custom property in an existing Excel file with Aspose.Cells .NET | Programmatically change Excel workbook metadata using Aspose.Cells in C# | Save custom document properties when creating a new workbook with Aspose.Cells
// Tags: Aspose.Cells custom document property update | C# set Excel workbook metadata | Aspose.Cells DocumentVersion property | modify Excel custom properties programmatically

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The code loads an existing Excel file (or creates a new one), checks for a custom document property named "DocumentVersion", sets its value to "2.0" or adds the property if it does not exist, and then saves the workbook. Exception handling ensures any errors are reported.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                Workbook workbook;

                // Load existing workbook if file exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Set or add a custom document property "DocumentVersion"
                const string propertyName = "DocumentVersion";
                var customProps = workbook.CustomDocumentProperties;

                if (customProps.Contains(propertyName))
                {
                    customProps[propertyName].Value = "2.0";
                }
                else
                {
                    customProps.Add(propertyName, "2.0");
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
