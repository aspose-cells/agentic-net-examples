// Title: Unit test for external link path change and Ribbon XML persistence in Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add an external link, modify its OriginalDataSource, assign custom Ribbon XML, save as XLSM to a MemoryStream, reload the file, and assert that both the updated link path and the Ribbon XML survive the round‑trip using Aspose.Cells for .NET.
// Keywords: Aspose.Cells unit test | external link path validation | RibbonXml persistence | XLSM custom ribbon | C# workbook round‑trip test | Aspose.Cells ExternalLink | Aspose.Cells RibbonXml
// Common Searches: Aspose.Cells unit test external link | persist custom ribbon xml in XLSM | verify external link OriginalDataSource after save | C# test for RibbonXml property | Aspose.Cells round‑trip validation
// Developer Intent: Create an automated test that confirms a modified external link URL and custom Ribbon XML are correctly saved and retrieved in an XLSM workbook using Aspose.Cells for .NET.
// Use Cases: Assert that ExternalLink.OriginalDataSource reflects a new URL after being changed. | Validate that the RibbonXml property is stored and loaded unchanged. | Ensure both the updated external link and custom ribbon survive a memory‑stream save/load cycle.
// AI Prompts: Generate an MSTest method that builds a Workbook, adds an external link, changes its path, sets RibbonXml, saves as Xlsm to a MemoryStream, reloads, and asserts persistence. | Provide NUnit test code for verifying external link path modification and custom ribbon integration with Aspose.Cells for .NET. | Explain how to mock external link validation while unit testing RibbonXml persistence in Aspose.Cells.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Demonstrates how to create a workbook, add an external link, modify its OriginalDataSource, assign custom Ribbon XML, save as XLSM to a MemoryStream, reload the file, and assert that both the updated link path and the Ribbon XML survive the round‑trip using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                ValidateExternalLinkPathChangeAndRibbonIntegration();
                Console.WriteLine("All validations passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        static void ValidateExternalLinkPathChangeAndRibbonIntegration()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add an external link to the workbook
            string originalPath = "https://example.com/oldpath/file.xlsx";
            string[] sheetNames = new string[] { "Sheet1!A1" };
            int linkIndex = workbook.Worksheets.ExternalLinks.Add(originalPath, sheetNames);
            ExternalLink link = workbook.Worksheets.ExternalLinks[linkIndex];

            // Verify that OriginalDataSource initially matches the added path
            Debug.Assert(originalPath == link.OriginalDataSource,
                "Initial OriginalDataSource does not match the added path.");

            // Modify the external link path
            string modifiedPath = originalPath.Replace("oldpath", "newpath");
            link.OriginalDataSource = modifiedPath;

            // Validate the path change
            Debug.Assert(modifiedPath == link.OriginalDataSource,
                "Modified OriginalDataSource was not set correctly.");

            // Set custom Ribbon XML (XLSM required for Ribbon XML)
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
            workbook.RibbonXml = ribbonXml;

            // Verify RibbonXml property
            Debug.Assert(ribbonXml == workbook.RibbonXml,
                "RibbonXml property was not set correctly.");

            // Save the workbook to a memory stream (XLSM required for Ribbon XML)
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsm);
                ms.Position = 0; // Reset stream position for reading

                // Load the workbook from the memory stream
                Workbook loadedWorkbook = new Workbook(ms);

                // Verify that the RibbonXml persisted
                Debug.Assert(ribbonXml == loadedWorkbook.RibbonXml,
                    "RibbonXml did not persist after saving and loading.");

                // Verify that the external link path persisted
                ExternalLink loadedLink = loadedWorkbook.Worksheets.ExternalLinks[0];
                Debug.Assert(modifiedPath == loadedLink.OriginalDataSource,
                    "Modified OriginalDataSource did not persist after saving and loading.");
            }
        }
    }
}
