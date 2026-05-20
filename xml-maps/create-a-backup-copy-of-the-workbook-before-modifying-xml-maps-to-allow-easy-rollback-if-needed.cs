using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XmlMapBackupDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the original workbook
            string originalPath = "OriginalWorkbook.xlsx";

            // Verify the original file exists
            if (!File.Exists(originalPath))
            {
                Console.WriteLine($"File not found: {originalPath}");
                return;
            }

            // Load the original workbook
            Workbook originalWorkbook = new Workbook(originalPath);

            // -----------------------------------------------------------------
            // Create a backup copy of the workbook before making any changes
            // -----------------------------------------------------------------
            // Create an empty workbook instance
            Workbook backupWorkbook = new Workbook();

            // Copy all contents from the original workbook to the backup workbook
            backupWorkbook.Copy(originalWorkbook);

            // Save the backup workbook to a separate file
            string backupPath = "BackupWorkbook.xlsx";
            backupWorkbook.Save(backupPath);
            Console.WriteLine($"Backup created at: {backupPath}");

            // ---------------------------------------------------------------
            // Now modify XML maps in the original workbook as needed
            // ---------------------------------------------------------------
            // Example: Add a new XML map to the original workbook
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item' type='xs:string'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                </xs:schema>";

            // Add the XML map and retrieve its index
            int mapIndex = originalWorkbook.Worksheets.XmlMaps.Add(xmlSchema);

            // Access the newly added XML map (optional, e.g., set a name)
            XmlMap xmlMap = originalWorkbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "SampleMap";

            // Save the modified workbook
            string modifiedPath = "ModifiedWorkbook.xlsx";
            originalWorkbook.Save(modifiedPath);
            Console.WriteLine($"Modified workbook saved at: {modifiedPath}");
        }
    }
}