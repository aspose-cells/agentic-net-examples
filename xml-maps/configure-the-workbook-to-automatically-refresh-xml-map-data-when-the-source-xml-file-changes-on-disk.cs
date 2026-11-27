// Title: C# example: Auto‑refresh an Aspose.Cells workbook XML map when the source XML file is modified
// Common Searches: how to use FileSystemWatcher to update an Aspose.Cells XML map in C# | auto refresh XML map data in Excel with Aspose.Cells when the XML file changes | C# code to monitor external XML file and reimport into an Aspose.Cells workbook | Aspose.Cells ImportXml method refresh on file change example | recalculate formulas after XML import in an Aspose.Cells workbook | synchronize Excel workbook with external XML using Aspose.Cells and FileSystemWatcher
// Tags: Aspose.Cells XML map refresh on file change | C# FileSystemWatcher integration with Aspose.Cells | ImportXml method auto update workbook | Excel workbook synchronization with external XML | CalculateFormula after XML import Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace XmlMapAutoRefreshDemo
{
    // The sample creates an Aspose.Cells Workbook, adds an XML map from an XSD schema, and imports XML data into the first worksheet. A FileSystemWatcher monitors the source XML file; when the file is modified, the program re‑imports the updated XML, optionally recalculates formulas, and saves the workbook, keeping the Excel file in sync with the external XML source.
    class Program
    {
        // Paths – adjust these to point to existing files on your machine
        private const string XmlSchemaPath = "sample.xsd";   // XSD that defines the XML map
        private const string XmlDataPath   = "data.xml";    // XML data file to be monitored
        private const string WorkbookPath  = "WorkbookWithXmlMap.xlsx";

        static void Main(string[] args)
        {
            // Verify that required files exist
            if (!File.Exists(XmlSchemaPath))
            {
                Console.WriteLine($"Schema file not found: {XmlSchemaPath}");
                return;
            }

            if (!File.Exists(XmlDataPath))
            {
                Console.WriteLine($"XML data file not found: {XmlDataPath}");
                return;
            }

            // -----------------------------------------------------------------
            // 1. Create a workbook and add an XML map based on the XSD schema
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();

            // Add the XML map – the Add method expects a path to an XSD or XML file
            int mapIndex = workbook.Worksheets.XmlMaps.Add(XmlSchemaPath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            Console.WriteLine($"Added XmlMap: Name = {xmlMap.Name}, Root = {xmlMap.RootElementName}");

            // Import the initial XML data into the first worksheet (Sheet1) starting at A1
            ImportXmlIntoWorkbook(workbook, XmlDataPath);
            workbook.Save(WorkbookPath);
            Console.WriteLine($"Workbook saved to {WorkbookPath}");

            // ---------------------------------------------------------------
            // 2. Set up a FileSystemWatcher to monitor changes to the XML file
            // ---------------------------------------------------------------
            string directory = Path.GetDirectoryName(Path.GetFullPath(XmlDataPath));
            string fileName  = Path.GetFileName(XmlDataPath);

            using (FileSystemWatcher watcher = new FileSystemWatcher(directory, fileName))
            {
                watcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.Attributes;
                watcher.Changed += (sender, e) => OnXmlFileChanged(workbook, e.FullPath);
                watcher.EnableRaisingEvents = true;

                Console.WriteLine($"Monitoring changes to '{XmlDataPath}'. Press ENTER to exit.");
                Console.ReadLine(); // Keep the application running
            }

            // Dispose workbook resources
            workbook.Dispose();
        }

        private static void ImportXmlIntoWorkbook(Workbook workbook, string xmlPath)
        {
            try
            {
                // Ensure the target sheet exists
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Import (or update) the XML data starting at cell A1 (row 0, column 0)
                workbook.ImportXml(xmlPath, sheet.Name, 0, 0);
                Console.WriteLine($"XML data imported from '{xmlPath}' at {DateTime.Now:HH:mm:ss}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error importing XML: {ex.Message}");
            }
        }

        /// <summary>
        /// Event handler invoked when the monitored XML file changes.
        /// Re-imports the XML data into the workbook and saves the workbook.
        /// </summary>
        private static void OnXmlFileChanged(Workbook workbook, string changedFilePath)
        {
            // FileSystemWatcher may raise multiple events; add a short debounce.
            System.Threading.Thread.Sleep(200); // 200 ms pause to allow the write operation to complete

            Console.WriteLine($"Detected change in XML file: {changedFilePath}");

            // Re-import the updated XML data
            ImportXmlIntoWorkbook(workbook, changedFilePath);

            // Optional: recalculate formulas if the workbook contains any
            try
            {
                workbook.CalculateFormula();
            }
            catch (Exception calcEx)
            {
                Console.WriteLine($"Formula calculation error: {calcEx.Message}");
            }

            // Save the workbook to persist the refreshed data
            try
            {
                workbook.Save(WorkbookPath);
                Console.WriteLine($"Workbook refreshed and saved to {WorkbookPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
    }
}
