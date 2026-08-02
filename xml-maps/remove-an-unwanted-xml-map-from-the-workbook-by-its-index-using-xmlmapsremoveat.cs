using System;
using Aspose.Cells;

class RemoveXmlMapDemo
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure there is at least one XML map before attempting removal
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            // Index of the XML map to remove (adjust as needed)
            int indexToRemove = 0;

            // Validate the index to avoid ArgumentOutOfRangeException
            if (indexToRemove >= 0 && indexToRemove < workbook.Worksheets.XmlMaps.Count)
            {
                // Remove the XML map at the specified index
                workbook.Worksheets.XmlMaps.RemoveAt(indexToRemove);
                Console.WriteLine($"Removed XML map at index {indexToRemove}.");
            }
            else
            {
                Console.WriteLine("Specified index is out of range.");
            }
        }
        else
        {
            Console.WriteLine("No XML maps found in the workbook.");
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
        Console.WriteLine("Workbook saved successfully.");
    }
}