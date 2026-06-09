using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsExamples
{
    public class CustomXmlPartNamingConvention
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Define the project identifier that will be prefixed to each custom XML part ID
                const string projectIdentifier = "Proj123_";

                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // new workbook

                // Example of loading an existing workbook (uncomment and provide a valid path)
                // string inputPath = "input.xlsx";
                // if (File.Exists(inputPath))
                // {
                //     workbook = new Workbook(inputPath);
                // }

                // Add a sample custom XML part to demonstrate the naming convention
                string sampleXml = "<root><item>Sample</item></root>";
                byte[] xmlBytes = Encoding.UTF8.GetBytes(sampleXml);
                int partIndex = workbook.CustomXmlParts.Add(xmlBytes, null);
                CustomXmlPart part = workbook.CustomXmlParts[partIndex];
                // Assign an initial ID (normally a GUID)
                part.ID = Guid.NewGuid().ToString();

                // Apply the naming convention: prefix each custom XML part ID with the project identifier
                foreach (CustomXmlPart xmlPart in workbook.CustomXmlParts)
                {
                    // Ensure we don't double‑prefix if the ID already starts with the identifier
                    if (!xmlPart.ID.StartsWith(projectIdentifier, StringComparison.Ordinal))
                    {
                        xmlPart.ID = projectIdentifier + xmlPart.ID;
                    }
                }

                // Save the workbook to verify the changes
                string outputPath = "CustomXmlPart_Named.xlsx";
                workbook.Save(outputPath);

                // Reload the workbook to demonstrate that the IDs were saved correctly
                if (File.Exists(outputPath))
                {
                    Workbook reloadedWorkbook = new Workbook(outputPath);
                    foreach (CustomXmlPart xmlPart in reloadedWorkbook.CustomXmlParts)
                    {
                        Console.WriteLine("Custom XML Part ID: " + xmlPart.ID);
                    }
                }
                else
                {
                    Console.WriteLine($"Failed to save workbook: {outputPath} not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}