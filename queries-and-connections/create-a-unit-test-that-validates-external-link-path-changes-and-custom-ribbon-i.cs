using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRunner
{
    class Program
    {
        private const string TempFolder = "TempTestFiles";

        static void Main()
        {
            GlobalSetup();

            try
            {
                ExternalLinkPathModification_ShouldUpdateOriginalDataSource();
                RibbonXmlIntegration_ShouldPersistCustomRibbon();
                Console.WriteLine("All tests passed.");
            }
            finally
            {
                GlobalTeardown();
            }
        }

        static void GlobalSetup()
        {
            if (Directory.Exists(TempFolder))
                Directory.Delete(TempFolder, true);
            Directory.CreateDirectory(TempFolder);
        }

        static void GlobalTeardown()
        {
            if (Directory.Exists(TempFolder))
                Directory.Delete(TempFolder, true);
        }

        static void ExternalLinkPathModification_ShouldUpdateOriginalDataSource()
        {
            Workbook workbook = new Workbook();

            string originalPath = @"https://example.com/oldfolder/source.xlsx";
            string[] sheetNames = new[] { "Sheet1!A1" };
            int linkIndex = workbook.Worksheets.ExternalLinks.Add(originalPath, sheetNames);

            if (workbook.Worksheets.ExternalLinks.Count != 1)
                throw new Exception("External link count mismatch.");

            ExternalLink link = workbook.Worksheets.ExternalLinks[linkIndex];
            if (link.OriginalDataSource != originalPath)
                throw new Exception("Original data source mismatch.");

            string modifiedPath = originalPath.Replace("oldfolder", "newfolder");
            link.OriginalDataSource = modifiedPath;

            if (link.OriginalDataSource != modifiedPath)
                throw new Exception("Modified path not set.");

            string filePath = Path.Combine(TempFolder, "ExternalLinkPathTest.xlsx");
            workbook.Save(filePath);

            Workbook reloaded = new Workbook(filePath);
            ExternalLink reloadedLink = reloaded.Worksheets.ExternalLinks[linkIndex];
            if (reloadedLink.OriginalDataSource != modifiedPath)
                throw new Exception("Modified path not persisted after reload.");
        }

        static void RibbonXmlIntegration_ShouldPersistCustomRibbon()
        {
            Workbook workbook = new Workbook();

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

            if (string.IsNullOrEmpty(workbook.RibbonXml) || !workbook.RibbonXml.Contains("customTab"))
                throw new Exception("RibbonXml not set correctly.");

            string filePath = Path.Combine(TempFolder, "RibbonXmlTest.xlsm");
            workbook.Save(filePath);

            Workbook reloaded = new Workbook(filePath);
            if (string.IsNullOrEmpty(reloaded.RibbonXml) || !reloaded.RibbonXml.Contains("customTab"))
                throw new Exception("RibbonXml not persisted after reload.");
        }
    }
}