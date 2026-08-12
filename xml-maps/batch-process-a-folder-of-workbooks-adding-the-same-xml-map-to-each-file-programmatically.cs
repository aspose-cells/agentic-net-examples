// Title: C# – Batch Add a Shared XML Map to Multiple Excel Workbooks with Aspose.Cells
// Description: A console utility that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, inserts the same XSD schema as an XML map using Worksheets.XmlMaps.Add, assigns a friendly name, overwrites the original file, and logs successes or errors.
// Keywords: Aspose.Cells | C# | .NET | XML map | XmlMaps.Add | batch process Excel | add XSD to workbook | automate Excel XML schema | folder scanning | Excel automation
// Common Searches: batch add xml map to excel workbooks c# | aspacells add same xml schema to multiple files | how to apply an XSD as an XML map to all workbooks in a folder | c# program to insert xml map into many Excel files | aspocells XmlMaps.Add example for batch processing
// Developer Intent: Insert an identical XML map into every workbook within a specified directory.
// Use Cases: Standardize a set of reporting templates with a common XML schema before data export. | Prepare bulk workbook files for downstream XML import pipelines. | Integrate XML map insertion into CI/CD workflows that generate Excel reports.
// AI Prompts: Write C# code that reads an XSD file and adds it as an XML map to all .xlsx files in a directory using Aspose.Cells. | Modify the batch script to recurse into subfolders and log the processing results to a CSV file. | Explain how to programmatically verify that the XML map was added to each workbook after the batch run.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchXmlMapAdder
{
    // A console utility that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, inserts the same XSD schema as an XML map using Worksheets.XmlMaps.Add, assigns a friendly name, overwrites the original file, and logs successes or errors.
    class Program
    {
        // Path to the folder containing the Excel workbooks
        private const string InputFolder = @"C:\Workbooks";

        // XML schema (XSD) content or file path to be added as a map.
        // Here we use a string containing the schema; you can also provide a file path.
        private const string XmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                            <xs:element name='Root'>
                                                <xs:complexType>
                                                    <xs:sequence>
                                                        <xs:element name='Item' type='xs:string'/>
                                                    </xs:sequence>
                                                </xs:complexType>
                                            </xs:element>
                                          </xs:schema>";

        static void Main()
        {
            // Validate that the input folder exists
            if (!Directory.Exists(InputFolder))
            {
                Console.WriteLine($"Folder not found: {InputFolder}");
                return;
            }

            // Get all Excel files in the folder (you can adjust the search pattern as needed)
            string[] excelFiles = Directory.GetFiles(InputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in excelFiles)
            {
                try
                {
                    // Load the workbook from the file
                    Workbook workbook = new Workbook(filePath);

                    // Add the XML map to the workbook (returns the index of the new map)
                    int mapIndex = workbook.Worksheets.XmlMaps.Add(XmlSchema);

                    // Optional: set a friendly name for the map
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                    xmlMap.Name = "SharedXmlMap";

                    // Save the workbook, overwriting the original file
                    workbook.Save(filePath);

                    Console.WriteLine($"Processed: {Path.GetFileName(filePath)} (Map added at index {mapIndex})");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
