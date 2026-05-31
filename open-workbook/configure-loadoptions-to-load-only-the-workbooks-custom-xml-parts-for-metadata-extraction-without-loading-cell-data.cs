using System;
using System.Text;
using Aspose.Cells;

class LoadCustomXmlOnly
{
    static void Main()
    {
        // Path to the workbook file
        string filePath = "input.xlsx";

        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Configure the LoadFilter to load only the workbook structure (no cell data)
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.Structure);

        // Optional: do not keep unparsed data to improve performance
        loadOptions.KeepUnparsedData = false;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Access and display custom XML parts
        Console.WriteLine($"Custom XML parts count: {workbook.CustomXmlParts.Count}");
        for (int i = 0; i < workbook.CustomXmlParts.Count; i++)
        {
            var part = workbook.CustomXmlParts[i];
            // Convert the XML data (byte array) to a string for display
            string xmlContent = Encoding.UTF8.GetString(part.Data);
            Console.WriteLine($"--- Custom XML Part {i + 1} ---");
            Console.WriteLine(xmlContent);
        }

        // No saving required as we only needed metadata extraction
    }
}