using System;
using Aspose.Cells;

class ListCombineFormats
{
    static void Main()
    {
        // List all loadable formats (these can be combined because they can be opened as a Workbook)
        Console.WriteLine("Loadable file formats that can be merged (based on LoadFormat enum):");
        foreach (LoadFormat fmt in Enum.GetValues(typeof(LoadFormat)))
        {
            if (fmt == LoadFormat.Unknown) continue; // skip unknown
            Console.WriteLine("- " + fmt);
        }

        // List all saveable formats (these can be the result of a merge)
        Console.WriteLine("\nSaveable file formats (based on SaveFormat enum):");
        foreach (SaveFormat fmt in Enum.GetValues(typeof(SaveFormat)))
        {
            if (fmt == SaveFormat.Unknown) continue; // skip unknown
            Console.WriteLine("- " + fmt);
        }

        // Demonstrate merging two different formats
        // Create a workbook in the default Xlsx format
        Workbook destWorkbook = new Workbook(); // default Xlsx
        destWorkbook.Worksheets[0].Cells["A1"].PutValue("Data from Xlsx");

        // Create a workbook from a CSV file (using the constructor that accepts FileFormatType)
        Workbook sourceWorkbook = new Workbook(FileFormatType.Csv);
        sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Data from CSV");

        // Combine the CSV workbook into the Xlsx workbook
        destWorkbook.Combine(sourceWorkbook);

        // Save the combined workbook (using the Save method)
        destWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}