using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class ExternalLinkAndRibbonDemo
    {
        // Sample ribbon XML used for the demo
        private const string SampleRibbonXml =
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

        // Original external link used in the workbook
        private const string OriginalLink =
            "https://arcusventures.sharepoint.com/Fund II/example.xlsx";

        // Expected link after replacement
        private const string ExpectedModifiedLink =
            "/sites/shared/shared documents/Fund II/example.xlsx";

        static void Main()
        {
            try
            {
                ExternalLinkPathModification_ShouldUpdateOriginalDataSource();
                RibbonXml_ShouldPersistAfterSaveAndLoad();
                Console.WriteLine("All operations completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Replicates the external‑link modification test
        private static void ExternalLinkPathModification_ShouldUpdateOriginalDataSource()
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Add a formula that references an external workbook via URL
            var ws = workbook.Worksheets[0];
            ws.Cells["A1"].Formula = $"='{{{OriginalLink}}}[Sheet1]'!A1";

            // Verify that an external link was created
            if (workbook.Worksheets.ExternalLinks.Count != 1)
                throw new InvalidOperationException("External link count should be 1.");

            // Modify the external link path using OriginalDataSource property
            var externalLink = workbook.Worksheets.ExternalLinks[0];
            string modifiedLink = externalLink.OriginalDataSource.Replace(
                "https://arcusventures.sharepoint.com/Fund II/",
                "/sites/shared/shared documents/Fund II/");

            externalLink.OriginalDataSource = modifiedLink;

            // Assert that the modification took effect
            if (externalLink.OriginalDataSource != ExpectedModifiedLink)
                throw new InvalidOperationException("External link path was not updated to the expected value.");
        }

        // Replicates the ribbon‑XML persistence test
        private static void RibbonXml_ShouldPersistAfterSaveAndLoad()
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Set custom ribbon XML
            workbook.RibbonXml = SampleRibbonXml;

            // Save workbook to a temporary file
            string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsm");
            workbook.Save(tempFile);

            // Ensure the file was created
            if (!File.Exists(tempFile))
                throw new FileNotFoundException("The workbook file was not saved correctly.", tempFile);

            // Load the workbook from the file
            var loadedWorkbook = new Workbook(tempFile);

            // Verify that the RibbonXml property persisted correctly
            if (loadedWorkbook.RibbonXml != SampleRibbonXml)
                throw new InvalidOperationException("RibbonXml did not persist after saving and loading the workbook.");

            // Clean up temporary file
            try
            {
                File.Delete(tempFile);
            }
            catch
            {
                // Ignored – cleanup failure should not affect program flow
            }
        }
    }
}