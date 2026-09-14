// Title: Export a custom Ribbon XML from an Aspose.Cells workbook to a separate .xml file using C#
// AI Prompts: Write C# code that creates a Workbook, assigns custom UI markup to the RibbonXml property, and saves the RibbonXml content to an external .xml file. | Show how to persist the RibbonXml of an Aspose.Cells workbook as a standalone XML file for source‑control integration in a .NET project. | Provide a C# example that both saves a macro‑enabled workbook and writes its custom ribbon definition to a separate file.
// Common Searches: how to save Aspose.Cells custom ribbon XML to a separate file in C# | export RibbonXml property to .xml for version control Aspose.Cells | C# example for writing workbook RibbonXml to external XML file | Aspose.Cells custom UI ribbon definition export to file | store Aspose.Cells ribbon customization in source control
// Tags: export RibbonXml Aspose.Cells C# | write custom ribbon XML to external file | Aspose.Cells workbook ribbon definition version control | save RibbonXml as separate .xml file | C# Aspose.Cells custom UI export

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates creating an Aspose.Cells Workbook, setting its RibbonXml property with custom UI markup, and exporting that XML to a separate .xml file for version‑control, while optionally saving the workbook.
    public class ExportRibbonXmlDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle create rule)
                Workbook workbook = new Workbook();

                // Define custom ribbon XML
                string ribbonXml =
                    "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                    "  <ribbon>" +
                    "    <tabs>" +
                    "      <tab id=\"customTab\" label=\"My Tab\">" +
                    "        <group id=\"customGroup\" label=\"My Group\">" +
                    "          <button id=\"customButton\" label=\"My Button\" size=\"large\" />" +
                    "        </group>" +
                    "      </tab>" +
                    "    </tabs>" +
                    "  </ribbon>" +
                    "</customUI>";

                // Set the RibbonXml property
                workbook.RibbonXml = ribbonXml;

                // Export the RibbonXml to a separate .xml file for version control
                string ribbonXmlPath = "RibbonCustom.xml";
                File.WriteAllText(ribbonXmlPath, workbook.RibbonXml);

                // Save the workbook (optional, demonstrates save rule)
                string workbookPath = "WorkbookWithRibbon.xlsm";
                workbook.Save(workbookPath);

                Console.WriteLine($"Ribbon XML exported to '{ribbonXmlPath}'.");
                Console.WriteLine($"Workbook saved to '{workbookPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportRibbonXmlDemo.Run();
        }
    }
}
