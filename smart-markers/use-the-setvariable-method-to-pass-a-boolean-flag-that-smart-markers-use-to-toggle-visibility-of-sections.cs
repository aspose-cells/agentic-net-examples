// Title: How to toggle visibility of Excel sections with Aspose.Cells Smart Markers by passing a Boolean flag via WorkbookDesigner.SetDataSource in C#
// Common Searches: Aspose.Cells C# set boolean variable for smart marker conditional sections | Using WorkbookDesigner.SetDataSource to hide or show rows in an Excel template | Smart markers conditional visibility based on a Boolean flag in .NET | Pass true false to smart markers to control section display in Aspose.Cells | Example of processing smart markers with a Boolean flag in C#
// Tags: WorkbookDesigner SetDataSource for smart markers | conditional smart marker visibility Aspose.Cells | toggle Excel sections with Boolean flag | C# smart marker processing example | Aspose.Cells template variable handling

using System;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerExample
{
    // The example loads a template workbook, creates a WorkbookDesigner, assigns a Boolean variable named "ShowSection" using SetDataSource, processes the smart markers so sections conditioned on the variable are shown or hidden, and saves the resulting file, with checks for missing template files and runtime error handling.
    class Program
    {
        static void Main(string[] args)
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: Template file \"{templatePath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook that contains Smart Markers
                Workbook workbook = new Workbook(templatePath);

                // Create a WorkbookDesigner for Smart Marker processing
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Pass a boolean flag to be used by Smart Markers for conditional visibility
                designer.SetDataSource("ShowSection", true);

                // Process the Smart Markers (this will evaluate the flag and hide/show sections accordingly)
                designer.Process();

                // Save the resulting workbook
                workbook.Save(resultPath);
                Console.WriteLine($"Processing completed. Result saved to \"{resultPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
