using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DisableDataTabDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Custom Ribbon XML that omits the default "Data" tab.
                string ribbonXml =
                    "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                    "  <ribbon>" +
                    "    <tabs>" +
                    "      <tab idMso=\"TabHome\" />" +
                    "      <tab idMso=\"TabInsert\" />" +
                    "      <tab idMso=\"TabPageLayout\" />" +
                    "      <tab idMso=\"TabFormulas\" />" +
                    "      <tab idMso=\"TabReview\" />" +
                    "      <tab idMso=\"TabView\" />" +
                    "    </tabs>" +
                    "  </ribbon>" +
                    "</customUI>";

                // Apply the custom Ribbon XML to the workbook
                workbook.RibbonXml = ribbonXml;

                string outputPath = "WorkbookWithoutDataTab.xlsm";

                // Save the workbook (macro-enabled format preserves Ribbon customizations)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            DisableDataTabDemo.Run();
        }
    }
}