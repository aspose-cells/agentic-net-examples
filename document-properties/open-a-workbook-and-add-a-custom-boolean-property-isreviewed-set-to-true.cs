// Title: Create and set a Boolean custom document property "IsReviewed" in an Excel file with Aspose.Cells (C#)
// Description: Shows how to load or generate a workbook, add a Boolean custom property named IsReviewed set to true via the CustomDocumentProperties collection, and save the updated .xlsx file.
// Keywords: Aspose.Cells | C# | Excel custom property | Boolean metadata | IsReviewed flag | Workbook.CustomDocumentProperties | add document property | save workbook | Excel file metadata
// Common Searches: Aspose.Cells C# add Boolean custom property | Set IsReviewed flag in Excel using Aspose | How to write custom document properties with Aspose.Cells | Create workbook and add metadata Aspose.Cells .NET | Read custom property from Excel Aspose.Cells
// Developer Intent: Insert a Boolean custom document property called IsReviewed with a true value into an Excel workbook and persist the change.
// Use Cases: Mark a report as reviewed before distribution. | Store processing status for automated pipelines. | Enable downstream tools to detect validated files. | Support compliance audits by flagging approved workbooks.
// AI Prompts: Generate C# code using Aspose.Cells that adds a Boolean custom property 'IsReviewed' to an existing workbook, includes error handling, and saves the file. | Provide an example that creates a new workbook, adds several custom properties of different types (Boolean, string, date), and writes the file to disk. | Explain how to retrieve the 'IsReviewed' custom property from a workbook with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to load or generate a workbook, add a Boolean custom property named IsReviewed set to true via the CustomDocumentProperties collection, and save the updated .xlsx file.
    public class AddCustomBooleanProperty
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook '{inputPath}': {ex.Message}");
                    return;
                }
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Add a custom Boolean property named "IsReviewed" with value true
            try
            {
                workbook.CustomDocumentProperties.Add("IsReviewed", true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to add custom property: {ex.Message}");
                return;
            }

            // Save the workbook with the new property
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
