// Title: Read and List ContentTypeProperties of an Excel Workbook using Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to open an existing Excel file with Aspose.Cells, retrieve the workbook's ContentTypeProperties collection, verify its presence, and iterate through each property to display Name, Value, Type, and IsNillable information in the console.
// Keywords: Aspose.Cells ContentTypeProperties | C# read Excel metadata | Workbook ContentTypePropertyCollection example | load Excel file Aspose.Cells | inspect custom content type in Excel
// Common Searches: how to get ContentTypeProperties from a workbook using Aspose.Cells | enumerate ContentTypeProperty objects in C# | check if Excel file contains content type metadata Aspose | list custom properties of an Excel workbook .NET | Aspose.Cells read custom content type information
// Developer Intent: Load an Excel workbook and output all its ContentTypeProperty metadata.
// Use Cases: Verify that required custom content‑type properties exist before further processing. | Debug missing or incorrect metadata by printing each property's details. | Create audit logs of workbook content‑type information for compliance reporting.
// AI Prompts: Write C# code to add a new ContentTypeProperty to a workbook with Aspose.Cells. | Show how to filter ContentTypeProperties where IsNillable is true and output only their names. | Provide an example of saving changes after modifying the ContentTypeProperties collection.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsContentTypeDemo
{
    // This example demonstrates how to open an existing Excel file with Aspose.Cells, retrieve the workbook's ContentTypeProperties collection, verify its presence, and iterate through each property to display Name, Value, Type, and IsNillable information in the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing Excel file.
            string filePath = "Sample.xlsx";

            // Load the workbook from the file using the provided constructor (load rule).
            using (Workbook workbook = new Workbook(filePath))
            {
                // Access the collection of ContentTypeProperty objects.
                ContentTypePropertyCollection contentProps = workbook.ContentTypeProperties;

                // If there are no content type properties, inform the user.
                if (contentProps.Count == 0)
                {
                    Console.WriteLine("The workbook does not contain any ContentTypeProperties.");
                }
                else
                {
                    // Iterate through each property and display its details.
                    for (int i = 0; i < contentProps.Count; i++)
                    {
                        ContentTypeProperty prop = contentProps[i];
                        Console.WriteLine($"Property #{i + 1}");
                        Console.WriteLine($"  Name       : {prop.Name}");
                        Console.WriteLine($"  Value      : {prop.Value}");
                        Console.WriteLine($"  Type       : {prop.Type}");
                        Console.WriteLine($"  IsNillable : {prop.IsNillable}");
                        Console.WriteLine();
                    }
                }
            }

            // Keep console window open if run outside an IDE.
            Console.WriteLine("Inspection completed. Press any key to exit.");
            Console.ReadKey();
        }
    }
}
