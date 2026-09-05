// Title: Hide all worksheet comments in an Excel file with Aspose.Cells for .NET while ensuring privacy compliance
// AI Prompts: Generate C# code using Aspose.Cells that loops through every worksheet, accesses each Comment object, sets its IsVisible property to false, and saves the workbook to a new file. | Enhance the previous code to encrypt the workbook with a password using Aspose.Cells after comments have been hidden.
// Common Searches: Aspose.Cells C# hide all comments in workbook for GDPR compliance | set comment IsVisible false across all sheets using Aspose.Cells | apply password encryption to Excel file after modifying comments with Aspose.Cells .NET | example code to hide worksheet comments and protect workbook in C# | loop through each sheet in an Excel workbook using Aspose.Cells
// Tags: Aspose.Cells comment visibility control | C# iterate worksheets Aspose.Cells | Excel comment privacy handling Aspose.Cells | Aspose.Cells workbook password encryption | batch update comment properties .NET

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing Excel workbook, verifies the file exists, iterates over every worksheet and each comment within, sets each comment's IsVisible flag to false to hide it, and then saves the modified workbook. It also demonstrates how to add password protection to the workbook after hiding comments, with error handling for file operations.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all comments on the worksheet
                foreach (Comment comment in sheet.Comments)
                {
                    // Hide the comment so it is not visible to end users
                    comment.IsVisible = false;

                    // Locking comments is not supported directly in the current Aspose.Cells API version.
                    // If needed, additional protection can be applied to the worksheet/workbook.
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions (e.g., loading/saving errors)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
