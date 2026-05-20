using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class FindMappedXmlCell
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Sample XML (not used for mapping in this demo)
                string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<Root>
    <Item id='123' name='SampleItem'>Value</Item>
</Root>";

                // Directly put the value that would come from the XML attribute into cell A1
                cells["A1"].PutValue("SampleItem");

                // Set up find options to search for the exact value in cell values
                FindOptions findOptions = new FindOptions
                {
                    LookInType = LookInType.Values,
                    LookAtType = LookAtType.EntireContent
                };

                // Locate the cell that holds the value
                Cell foundCell = cells.Find("SampleItem", null, findOptions);

                if (foundCell != null)
                {
                    Console.WriteLine($"Value found at cell: {foundCell.Name}");
                }
                else
                {
                    Console.WriteLine("Value not found.");
                }

                // Save the workbook (ensure the directory exists)
                string outputPath = "MappedXmlFindDemo.xlsx";

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}