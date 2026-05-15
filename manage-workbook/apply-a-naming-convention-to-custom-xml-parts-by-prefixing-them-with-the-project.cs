using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsExamples
{
    public class CustomXmlPartNamingConventionDemo
    {
        public static void Run()
        {
            // Define the project identifier that will be used as a prefix
            string projectId = "ProjA";

            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Prepare sample XML data for the custom XML part
            string xmlData = "<root><item>Sample Data</item></root>";
            byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlData);

            // Add the custom XML part to the workbook (no schema in this example)
            int partIndex = workbook.CustomXmlParts.Add(xmlBytes, null);

            // Retrieve the added part
            CustomXmlPart part = workbook.CustomXmlParts[partIndex];

            // Generate a new GUID for the part's original ID
            string originalId = Guid.NewGuid().ToString();

            // Apply the naming convention by prefixing the project identifier
            part.ID = $"{projectId}_{originalId}";

            // Save the workbook (save rule)
            string outputPath = "CustomXmlPart_Named.xlsx";
            workbook.Save(outputPath);

            // Load the workbook back (load rule) to verify the ID
            Workbook loadedWorkbook = new Workbook(outputPath);
            CustomXmlPart loadedPart = loadedWorkbook.CustomXmlParts.SelectByID(part.ID);

            // Output the ID to confirm the naming convention was applied
            Console.WriteLine("Original ID (without prefix): " + originalId);
            Console.WriteLine("Prefixed ID stored in workbook: " + loadedPart?.ID);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CustomXmlPartNamingConventionDemo.Run();
        }
    }
}