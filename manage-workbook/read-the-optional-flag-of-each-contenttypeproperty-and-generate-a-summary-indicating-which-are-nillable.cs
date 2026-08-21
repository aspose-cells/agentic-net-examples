// Title: List IsNillable Flags of ContentTypeProperty Objects in Aspose.Cells for .NET
// Description: Shows how to add ContentTypeProperty entries to a Workbook, set their IsNillable flag, iterate through all properties, output a console summary of each property's name and nillable status, and save the workbook.
// Keywords: Aspose.Cells | ContentTypeProperty | IsNillable | .NET | C# | metadata | nillable flag | list properties | workbook metadata
// Common Searches: Aspose.Cells check IsNillable for ContentTypeProperty | C# list workbook content type properties nillable | How to read optional flag of ContentTypeProperty in Aspose.Cells | Iterate ContentTypeProperties and display IsNillable | Save workbook after inspecting metadata Aspose.Cells
// Developer Intent: Read each ContentTypeProperty's IsNillable flag in a workbook and generate a readable console summary.
// Use Cases: Validate that required metadata fields are correctly marked as non‑nillable before publishing a workbook. | Create documentation that lists which workbook metadata properties can accept null values. | Log the nillable status of all ContentTypeProperties for compliance or auditing purposes.
// AI Prompts: Write C# code using Aspose.Cells to enumerate all ContentTypeProperty names with their IsNillable values and write the result to a text file. | Provide an example that filters only nillable ContentTypeProperties and exports their names to a CSV file. | Explain how to programmatically set IsNillable for multiple ContentTypeProperties based on a custom business rule.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    // Shows how to add ContentTypeProperty entries to a Workbook, set their IsNillable flag, iterate through all properties, output a console summary of each property's name and nillable status, and save the workbook.
    public class ContentTypePropertyNillableSummary
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook instance
                Workbook workbook = new Workbook();

                // Add several content type properties with different IsNillable settings
                int idx1 = workbook.ContentTypeProperties.Add("PropertyA", "ValueA", "text");
                workbook.ContentTypeProperties[idx1].IsNillable = true;

                int idx2 = workbook.ContentTypeProperties.Add("PropertyB", "2023-01-01T00:00:00", "DateTime");
                workbook.ContentTypeProperties[idx2].IsNillable = false;

                int idx3 = workbook.ContentTypeProperties.Add("PropertyC", "123", "number");
                workbook.ContentTypeProperties[idx3].IsNillable = true;

                // Build a summary of nillable flags
                StringBuilder summary = new StringBuilder();
                summary.AppendLine("ContentTypeProperty Nillable Summary:");
                for (int i = 0; i < workbook.ContentTypeProperties.Count; i++)
                {
                    ContentTypeProperty prop = workbook.ContentTypeProperties[i];
                    summary.AppendLine($"- Name: {prop.Name}, IsNillable: {prop.IsNillable}");
                }

                // Output the summary to the console
                Console.WriteLine(summary.ToString());

                // Save the workbook (optional, demonstrates lifecycle usage)
                workbook.Save("ContentTypePropertyNillableSummary.xlsx");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
