using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapRemoval
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the collection of XML maps in the workbook
            XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

            // Ensure there is at least one XML map to remove
            if (xmlMaps.Count > 0)
            {
                // Specify the index of the XML map to remove (e.g., the first map)
                int indexToRemove = 0;

                // Remove the XML map at the specified index
                xmlMaps.RemoveAt(indexToRemove);

                Console.WriteLine($"Removed XML map at index {indexToRemove}. Remaining maps: {xmlMaps.Count}");
            }
            else
            {
                Console.WriteLine("No XML maps found in the workbook.");
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}