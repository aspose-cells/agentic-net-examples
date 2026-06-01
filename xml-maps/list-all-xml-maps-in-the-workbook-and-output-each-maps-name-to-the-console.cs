using System;
using Aspose.Cells;

class ListXmlMaps
{
    static void Main()
    {
        // Load the workbook (provide the correct file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the collection of XML maps in the workbook
        XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

        // Iterate through each XML map and output its name
        for (int i = 0; i < xmlMaps.Count; i++)
        {
            XmlMap map = xmlMaps[i];
            Console.WriteLine(map.Name);
        }
    }
}