using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DisableDataTabInRibbonDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Custom Ribbon XML that defines the visible tabs.
            // The default "Data" tab is omitted, effectively disabling it.
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                // Include only the desired built‑in tabs (Home, Insert, Page Layout, Formulas, Review, View)
                "      <tab idMso=\"TabHome\" visible=\"true\" />" +
                "      <tab idMso=\"TabInsert\" visible=\"true\" />" +
                "      <tab idMso=\"TabPageLayout\" visible=\"true\" />" +
                "      <tab idMso=\"TabFormulas\" visible=\"true\" />" +
                "      <tab idMso=\"TabReview\" visible=\"true\" />" +
                "      <tab idMso=\"TabView\" visible=\"true\" />" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Apply the custom Ribbon XML to the workbook (property set)
            workbook.RibbonXml = ribbonXml;

            // Save the workbook (lifecycle: save). Use .xlsm because Ribbon customizations require a macro‑enabled format.
            workbook.Save("WorkbookWithoutDataTab.xlsm");
        }
    }
}