// Title: Disable the LinksUpToDate built‑in document property in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, sets the built‑in document property "LinksUpToDate" to false, and saves the workbook. | Show how to turn off automatic link validation in an Excel file by updating the LinksUpToDate property via the BuiltInDocumentProperties collection in Aspose.Cells.
// Common Searches: Aspose.Cells C# disable automatic link update in Excel workbook | set LinksUpToDate property to false using Aspose.Cells .NET | prevent external link checks when saving Excel file with Aspose.Cells | how to modify built‑in document properties in Aspose.Cells C#
// Tags: disable LinksUpToDate property Aspose.Cells | modify built-in document properties C# | turn off link validation Excel Aspose.Cells | set workbook property false Aspose.Cells | Aspose.Cells built-in properties manipulation

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, checks for the built‑in "LinksUpToDate" property, sets its value to false to stop link validation, and saves the modified file, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Disable the built‑in property "LinksUpToDate" to prevent link checks
            // The DocumentProperty class exposes a Value property for assignment.
            if (workbook.BuiltInDocumentProperties["LinksUpToDate"] != null)
            {
                workbook.BuiltInDocumentProperties["LinksUpToDate"].Value = false;
            }

            // Save the workbook (overwrites if the file already exists)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
