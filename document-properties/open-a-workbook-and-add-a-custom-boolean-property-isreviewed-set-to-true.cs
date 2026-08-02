using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertyDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing workbook to open
            string inputPath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Add a custom Boolean property named "IsReviewed" with value true
            workbook.CustomDocumentProperties.Add("IsReviewed", true);

            // Save the workbook with the new property
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            // Optional: Verify that the property was added
            Console.WriteLine($"Custom Property 'IsReviewed' Value: {workbook.CustomDocumentProperties["IsReviewed"].Value}");
        }
    }
}