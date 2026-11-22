// Title: Create a formatted address line from street, city, and ZIP cells using Aspose.Cells for .NET (C#)
// Common Searches: how to concatenate street, city and zip cell values into one address string with Aspose.Cells C# | Aspose.Cells retrieve cell string values and build formatted address line in .NET | write combined address to another cell and save workbook using Aspose.Cells | C# example merging multiple Excel cells into a single formatted address with Aspose.Cells | generate address line from separate Excel cells using Aspose.Cells API
// Tags: concatenate cell values Aspose.Cells C# | format address line Excel .NET | retrieve string value Aspose.Cells | write combined address to cell Aspose | save workbook Aspose.Cells .NET | populate Excel cells with address components C#

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example creates a new workbook, places street, city, and ZIP values in cells A1, B1, and C1, reads each cell's string value, concatenates them into a formatted address line ("street, city ZIP"), writes the result to cell D1, and saves the file as Address.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Example data: street in A1, city in B1, zip in C1
                var streetCell = sheet.Cells["A1"];
                var cityCell = sheet.Cells["B1"];
                var zipCell = sheet.Cells["C1"];

                // Populate cells
                streetCell.PutValue("123 Main St");
                cityCell.PutValue("Springfield");
                zipCell.PutValue("12345");

                // Retrieve formatted string values directly from the cells
                string street = streetCell.StringValue;
                string city = cityCell.StringValue;
                string zip = zipCell.StringValue;

                // Combine into a single formatted address line
                string addressLine = $"{street}, {city} {zip}";

                // Write the combined address to another cell (e.g., D1)
                sheet.Cells["D1"].PutValue(addressLine);

                // Save the workbook
                string outputPath = "Address.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
