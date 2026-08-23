// Title: Create a C# unit test to verify external link URL changes and custom Ribbon XML persistence in an Aspose.Cells workbook
// AI Prompts: Write a C# unit‑test method (using MSTest, NUnit, or xUnit) that adds an external link to a workbook, replaces a segment of its URL, assigns a custom Ribbon XML string, saves the workbook as .xlsm, reloads it, and asserts that both the modified link path and the RibbonXml property are retained with Aspose.Cells. | Extend the test to loop through multiple external links, update each URL fragment, and confirm that all links together with the custom ribbon configuration survive a save‑load cycle.
// Common Searches: how to unit test external link path modification in Aspose.Cells .NET | verify RibbonXml property persists after saving an Aspose.Cells workbook | c# test for updating external data source URLs and custom ribbon integration with Aspose.Cells
// Tags: Aspose.Cells external link URL unit testing | RibbonXml persistence check in Aspose.Cells | C# workbook save/load external data source verification | custom UI ribbon integration Aspose.Cells | unit test for external data source path change

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example defines a C# unit test that creates a workbook, adds an external link, modifies a portion of its URL, sets custom Ribbon XML, saves the file as .xlsm, reloads it, and uses assertions to ensure both the updated external link path and the RibbonXml setting are correctly persisted.
    public class ExternalLinkAndRibbonTests
    {
        // Sample ribbon XML used for the test
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

        public static void Main()
        {
            try
            {
                var test = new ExternalLinkAndRibbonTests();
                test.ValidateExternalLinkPathChangeAndRibbonIntegration();
                Console.WriteLine("Test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        public void ValidateExternalLinkPathChangeAndRibbonIntegration()
        {
            // ---------- Create a workbook and add an external link ----------
            Workbook workbook = new Workbook(); // create new workbook
            Worksheet sheet = workbook.Worksheets[0];
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Original external link path
            string originalPath = @"https://example.com/oldpath/source.xlsx";

            // Add the external link (file name and sheet reference)
            externalLinks.Add(originalPath, new[] { "Sheet1!A1" });

            // Set a formula that references the external link
            sheet.Cells["A1"].Formula = $"='https://example.com/oldpath/source.xlsx'!Sheet1!A1";

            // Verify that the external link was added
            AssertEqual(1, externalLinks.Count, "External link should be added.");

            // ---------- Modify the external link path ----------
            string modifiedPathFragment = "newpath";
            foreach (ExternalLink link in externalLinks)
            {
                // Replace the old path segment with the new one
                string modified = link.OriginalDataSource.Replace("oldpath", modifiedPathFragment);
                link.OriginalDataSource = modified;
            }

            // Verify that the path was updated
            foreach (ExternalLink link in externalLinks)
            {
                Assert(link.OriginalDataSource.Contains(modifiedPathFragment),
                    "External link path should contain the modified fragment.");
            }

            // ---------- Set custom Ribbon XML ----------
            workbook.RibbonXml = SampleRibbonXml;

            // Verify RibbonXml property is set correctly
            AssertEqual(SampleRibbonXml, workbook.RibbonXml, "RibbonXml should be set to the sample XML.");

            // ---------- Save the workbook ----------
            string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsm");
            try
            {
                workbook.Save(tempFile); // save workbook

                // Ensure the file was created before loading
                if (!File.Exists(tempFile))
                {
                    throw new FileNotFoundException("Saved workbook file not found.", tempFile);
                }

                // ---------- Load the workbook and re-validate ----------
                Workbook loadedWorkbook = new Workbook(tempFile); // load workbook

                // Verify RibbonXml persisted
                AssertEqual(SampleRibbonXml, loadedWorkbook.RibbonXml,
                    "Loaded workbook should retain the RibbonXml.");

                // Verify external link path persisted
                ExternalLinkCollection loadedLinks = loadedWorkbook.Worksheets.ExternalLinks;
                AssertEqual(1, loadedLinks.Count, "Loaded workbook should have one external link.");
                Assert(loadedLinks[0].OriginalDataSource.Contains(modifiedPathFragment),
                    "Loaded external link should contain the modified path fragment.");
            }
            finally
            {
                // Clean up temporary file
                if (File.Exists(tempFile))
                {
                    try
                    {
                        File.Delete(tempFile);
                    }
                    catch
                    {
                        // Ignored - cleanup failure should not affect test outcome
                    }
                }
            }
        }

        // Simple assertion helpers
        private void Assert(bool condition, string message)
        {
            if (!condition)
                throw new InvalidOperationException(message);
        }

        private void AssertEqual<T>(T expected, T actual, string message)
        {
            if (!Equals(expected, actual))
                throw new InvalidOperationException(message);
        }
    }
}
