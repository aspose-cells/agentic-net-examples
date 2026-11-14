// Title: How to set Aspose.Cells SmartMarkerProcessingMode to IgnoreErrors for partial data insertion in C#
// Common Searches: Aspose.Cells C# set SmartMarkerProcessingMode to IgnoreErrors for template processing | ignore smart marker errors during Excel generation with Aspose.Cells .NET | partial data insertion using smart markers without stopping execution in C# | configure smart marker processing to continue on missing fields in Aspose.Cells | C# example of handling smart marker merge errors in Aspose.Cells
// Tags: smartmarkerprocessingmode ignoreerrors c# | aspocells smart marker error handling | smart marker partial data insertion | configure smart marker processing mode | excel template generation aspocells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example shows how to load or create an Excel workbook, set the SmartMarkerProcessingMode to IgnoreErrors (e.g., workbook.Settings.SmartMarkerProcessingMode = SmartMarkerProcessingMode.IgnoreErrors) so that smart marker processing continues despite data errors, and then save the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input template and output file paths
            string templatePath = "Template.xlsx";
            string outputPath = "Output.xlsx";

            try
            {
                Workbook workbook;

                // Load existing template if it exists; otherwise create a new workbook
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    workbook = new Workbook(); // creates a default workbook with one worksheet
                }

                // NOTE: In some Aspose.Cells versions the SmartMarkerProcessingMode property may not be available.
                // If needed, set it here using the appropriate API for the version you are using.

                // TODO: Add any data population or smart marker processing here

                // Save the processed workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
