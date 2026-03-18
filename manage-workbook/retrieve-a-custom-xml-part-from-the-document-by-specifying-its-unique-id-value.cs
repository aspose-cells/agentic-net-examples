using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsCustomXmlPartDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook that contains custom XML parts
            string inputPath = "input.xlsx";

            // Load the workbook (load rule)
            Workbook workbook = new Workbook(inputPath);

            // Unique ID of the custom XML part to retrieve
            string partId = "2F087CB2-7CA8-43DA-B048-2E2F61F4936F";

            // Retrieve the custom XML part by its ID using SelectByID method
            CustomXmlPart customPart = workbook.CustomXmlParts.SelectByID(partId);

            if (customPart != null)
            {
                // Output the ID of the retrieved part
                Console.WriteLine("Custom XML Part ID: " + customPart.ID);

                // Convert the XML data (byte array) to a string for display
                string xmlContent = Encoding.UTF8.GetString(customPart.Data);
                Console.WriteLine("XML Content:");
                Console.WriteLine(xmlContent);
            }
            else
            {
                Console.WriteLine($"Custom XML Part with ID '{partId}' not found.");
            }

            // Save the workbook (save rule) – here we simply save to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
    }
}