using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsExamples
{
    public class CustomXmlPartNamingConventionDemo
    {
        // Project identifier to be prefixed to each custom XML part ID
        private const string ProjectId = "Proj123_";

        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Prepare sample XML data (no schema needed for this demo)
            string xmlData = "<root><item>Sample Data</item></root>";
            byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlData);

            // Add the custom XML part to the workbook (lifecycle: add)
            int partIndex = workbook.CustomXmlParts.Add(xmlBytes, null);
            CustomXmlPart part = workbook.CustomXmlParts[partIndex];

            // Generate a GUID and apply the naming convention (prefix with project ID)
            string originalGuid = Guid.NewGuid().ToString();
            part.ID = ProjectId + originalGuid; // e.g., "Proj123_2F087CB2-7CA8-43DA-B048-2E2F61F4936F"

            // Save the workbook (lifecycle: save)
            string outputPath = "CustomXmlPart_Named.xlsx";
            workbook.Save(outputPath);

            // Verify the file exists before loading
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("Saved workbook not found.", outputPath);

            // Load the workbook back (lifecycle: load) to verify the ID
            Workbook loadedWorkbook = new Workbook(outputPath);
            CustomXmlPart loadedPart = loadedWorkbook.CustomXmlParts.SelectByID(part.ID);

            // Output the ID to confirm the naming convention was applied
            Console.WriteLine("Original GUID: " + originalGuid);
            Console.WriteLine("Prefixed ID stored in workbook: " + (loadedPart != null ? loadedPart.ID : "Not found"));
        }
    }
}