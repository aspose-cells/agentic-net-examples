// Title: How to save a modified Excel workbook that includes custom document properties and XML parts as a new .xlsx file using Aspose.Cells for .NET
// AI Prompts: Load an existing .xlsx file, insert a custom document property and a custom XML part, then save the workbook as a new .xlsx with Aspose.Cells. | Add robust error handling to verify the source file exists before loading and to catch exceptions during the save operation using Aspose.Cells. | Adapt the code to keep the added custom properties when converting the workbook to another format such as .xlsb.
// Common Searches: asp.net set a custom document property in an Excel workbook using Aspose.Cells | how to embed an XML part into an Excel workbook with Aspose.Cells C# | preserve custom XML parts when saving a workbook using Aspose.Cells .NET | save modified workbook to a different file name while keeping custom properties Aspose.Cells | example code for validating source file existence before loading workbook Aspose.Cells
// Tags: custom document property insertion Aspose.Cells .NET | embed custom XML part Aspose.Cells | save workbook as new .xlsx Aspose.Cells | preserve custom properties during workbook save | source file validation Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Loads source.xlsx, adds a custom document property and a custom XML part, then saves the workbook as modified_output.xlsx, preserving the added metadata.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "modified_output.xlsx";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: The source file \"{sourcePath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(sourcePath);

            // -------------------------------------------------
            // Perform modifications here (e.g., add data, styles, etc.)
            // -------------------------------------------------

            // Example: add a custom document property
            workbook.CustomDocumentProperties.Add("ProjectName", "Aspose Integration");

            // Example: add a custom XML part (using byte[] overload)
            string xmlContent = "<root><item>Value</item></root>";
            byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlContent);
            // schemaData is optional; pass null if not required
            workbook.CustomXmlParts.Add(xmlBytes, null);

            // -------------------------------------------------
            // Save the modified workbook to a new .xlsx file.
            // All added custom properties and XML parts are preserved automatically.
            // -------------------------------------------------
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
