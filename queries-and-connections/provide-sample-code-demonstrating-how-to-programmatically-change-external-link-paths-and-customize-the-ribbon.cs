// Title: How to programmatically replace external link URLs and add custom Ribbon XML in an Excel .xlsm workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates over Workbook.Worksheets.ExternalLinks, replaces the old SharePoint base URL with a new one by updating OriginalDataSource and DataSource properties using Aspose.Cells. | Demonstrate how to assign a custom Ribbon XML string to Workbook.RibbonXml and save the workbook as a macro‑enabled .xlsm file with Aspose.Cells.
// Common Searches: Aspose.Cells replace external link URL SharePoint C# | set custom ribbon tab in Excel workbook using Aspose.Cells .NET | update external links collection programmatically Aspose.Cells example | save workbook with RibbonXml property macro enabled file Aspose.Cells | change data source path for external links in Excel via Aspose.Cells
// Tags: external link path substitution Aspose.Cells | ribbon customization via RibbonXml | macro-enabled workbook saving Aspose.Cells | sharepoint link path change Aspose.Cells | iterate external links collection C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, adds a dummy external link, loops through the ExternalLinks collection to replace an old SharePoint URL with a new one (updating both OriginalDataSource and DataSource), defines custom Ribbon XML, assigns it through the RibbonXml property, saves the file as a macro‑enabled .xlsm workbook, and prints the updated link information and ribbon status to the console.
    public class ExternalLinkAndRibbonDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Add a sample external link to demonstrate change
                // -------------------------------------------------
                // The external link points to a dummy file; in real scenarios the workbook would already contain links.
                workbook.Worksheets.ExternalLinks.Add(
                    "https://oldsharepoint.com/Projects/Report.xlsx",
                    new string[] { "Sheet1" });

                // -------------------------------------------------
                // 2. Change external link paths programmatically
                // -------------------------------------------------
                // Iterate through all external links and replace the old base URL with a new one.
                for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
                {
                    ExternalLink link = workbook.Worksheets.ExternalLinks[i];

                    // Use the OriginalDataSource property to get the stored path.
                    string originalPath = link.OriginalDataSource;

                    // Example replacement: change SharePoint domain and folder structure.
                    string updatedPath = originalPath.Replace(
                        "https://oldsharepoint.com/Projects/",
                        "https://newsharepoint.com/SharedDocs/");

                    // Assign the modified path back.
                    link.OriginalDataSource = updatedPath;

                    // Optional: also update DataSource if you want the active reference to change.
                    link.DataSource = updatedPath;
                }

                // -------------------------------------------------
                // 3. Customize the Ribbon UI via RibbonXml property
                // -------------------------------------------------
                string ribbonXml =
                    "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                    "  <ribbon>" +
                    "    <tabs>" +
                    "      <tab id=\"customTab\" label=\"My Custom Tab\">" +
                    "        <group id=\"customGroup\" label=\"My Group\">" +
                    "          <button id=\"customButton\" label=\"My Button\" size=\"large\" />" +
                    "        </group>" +
                    "      </tab>" +
                    "    </tabs>" +
                    "  </ribbon>" +
                    "</customUI>";

                // Assign the custom Ribbon XML to the workbook.
                workbook.RibbonXml = ribbonXml;

                // -------------------------------------------------
                // 4. Save the workbook (as a macro-enabled file to retain Ribbon customizations)
                // -------------------------------------------------
                string outputPath = "ExternalLinkAndRibbonDemo.xlsm";

                // Ensure the directory exists (in case a relative path is used)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);

                // -------------------------------------------------
                // 5. Verify changes (output to console)
                // -------------------------------------------------
                Console.WriteLine("External links after modification:");
                foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
                {
                    Console.WriteLine($"- OriginalDataSource: {link.OriginalDataSource}");
                    Console.WriteLine($"  DataSource: {link.DataSource}");
                }

                Console.WriteLine("Ribbon XML has been set: " + (workbook.RibbonXml != null));
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExternalLinkAndRibbonDemo.Run();
        }
    }
}
