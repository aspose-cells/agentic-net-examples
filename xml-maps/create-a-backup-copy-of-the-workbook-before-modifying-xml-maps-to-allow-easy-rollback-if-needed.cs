using System;
using Aspose.Cells;

namespace AsposeCellsBackupExample
{
    class Program
    {
        static void Main()
        {
            // Load the original workbook
            Workbook originalWorkbook = new Workbook("InputWorkbook.xlsx");

            // Create a backup copy of the workbook
            Workbook backupWorkbook = new Workbook();
            // Copy all contents from the original workbook to the backup workbook
            originalWorkbook.Copy(backupWorkbook);
            // Save the backup workbook to a separate file
            backupWorkbook.Save("BackupWorkbook.xlsx");

            // ----- Begin modifications to XML maps -----
            // Add a new XML map to the original workbook (example schema file path)
            int mapIndex = originalWorkbook.Worksheets.XmlMaps.Add("SampleSchema.xsd");
            // Access the newly added XML map
            XmlMap xmlMap = originalWorkbook.Worksheets.XmlMaps[mapIndex];
            // Optionally set a name for the map
            xmlMap.Name = "SampleMap";
            // ----- End modifications -----

            // Save the modified original workbook
            originalWorkbook.Save("ModifiedWorkbook.xlsx");
        }
    }
}