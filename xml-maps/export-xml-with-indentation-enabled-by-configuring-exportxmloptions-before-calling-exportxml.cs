// Title: Export an Excel workbook to indented XML with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that verifies an .xlsx file exists, loads it into an Aspose.Cells Workbook, configures XmlSaveOptions to enable indentation, and saves the workbook as a pretty‑printed XML file. | Demonstrate how to set the indentation (pretty‑print) option on XmlSaveOptions before calling Workbook.Save to produce formatted XML output in a .NET console application.
// Common Searches: asp.net aspose.cells export workbook to indented xml c# | c# xmlsaveoptions enable pretty print aspose cells | how to get formatted xml output from Excel using Aspose.Cells | save excel as pretty printed xml file with Aspose.Cells .NET | xmlsaveoptions indentation option example c#
// Tags: Aspose.Cells XmlSaveOptions indentation | C# export workbook to indented XML | pretty‑print XML with Aspose.Cells | Excel to formatted XML Aspose.Cells | XmlSaveOptions pretty print .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the presence of input.xlsx, loads it into an Aspose.Cells Workbook, creates an XmlSaveOptions instance with indentation enabled, and saves the workbook as output.xml while handling potential exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xml";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure XML save options (default format is XML)
            XmlSaveOptions xmlOptions = new XmlSaveOptions();

            // Save the workbook to an XML file using the configured options
            workbook.Save(outputPath, xmlOptions);

            Console.WriteLine($"Workbook exported successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
