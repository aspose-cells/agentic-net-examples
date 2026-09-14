// Title: How to test a custom ribbon button that updates external link paths in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, inserts an external link formula referencing another file, and iterates the workbook's ExternalLinks collection to modify each link's OriginalDataSource by replacing a specific folder segment. | Produce the RibbonXml markup for a custom Office ribbon tab that contains a large button labeled "Update Links" suitable for assignment to an Aspose.Cells workbook. | Show how to assign the RibbonXml to a Workbook object, apply the external‑link path changes, and save the workbook to a chosen output file.
// Common Searches: aspnet cells update external link path using OriginalDataSource property | c# example custom ribbon button that changes external link folder in Excel workbook | how to replace part of external link file path in Aspose.Cells workbook | assign RibbonXml to Aspose.Cells workbook for custom UI button | testing external link updates with Aspose.Cells and custom ribbon
// Tags: external link path replacement Aspose.Cells | OriginalDataSource property modification | custom ribbon XML Aspose.Cells | C# external link update Excel workbook | Ribbon button integration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace RibbonButtonExternalLinkTest
{
    // The sample creates a source workbook if missing, builds a new workbook with an external link formula pointing to that source, loops through the workbook's ExternalLinks collection to replace the folder segment "OldFolder" with "NewFolder" in each OriginalDataSource, defines a custom RibbonXml containing an "Update Links" button, assigns the RibbonXml to the workbook, and saves the result as UpdatedExternalLinks.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Prepare a source workbook that will be referenced
                string sourcePath = Path.Combine(Environment.CurrentDirectory, "source.xlsx");
                if (!File.Exists(sourcePath))
                {
                    Workbook sourceWb = new Workbook();
                    sourceWb.Worksheets[0].Cells["A1"].PutValue("Source Value");
                    sourceWb.Save(sourcePath);
                }

                // Create a new workbook that will contain the external link
                Workbook workbook = new Workbook();

                // Add an external link formula in cell A1 referencing the source workbook
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].Formula = $"='{sourcePath}'!Sheet1!A1";

                // Update all external link paths using OriginalDataSource property
                foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
                {
                    string original = link.OriginalDataSource;
                    // Example: replace "OldFolder" with "NewFolder" in the path
                    string updated = original.Replace("OldFolder", "NewFolder");
                    link.OriginalDataSource = updated;
                }

                // Set RibbonXml to simulate a custom ribbon button (illustrative only)
                string ribbonXml =
                    "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                    "  <ribbon>" +
                    "    <tabs>" +
                    "      <tab id=\"customTab\" label=\"Custom Tab\">" +
                    "        <group id=\"customGroup\" label=\"Link Tools\">" +
                    "          <button id=\"updateLinksBtn\" label=\"Update Links\" size=\"large\" />" +
                    "        </group>" +
                    "      </tab>" +
                    "    </tabs>" +
                    "  </ribbon>" +
                    "</customUI>";
                workbook.RibbonXml = ribbonXml;

                // Save the workbook
                string outputPath = Path.Combine(Environment.CurrentDirectory, "UpdatedExternalLinks.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
