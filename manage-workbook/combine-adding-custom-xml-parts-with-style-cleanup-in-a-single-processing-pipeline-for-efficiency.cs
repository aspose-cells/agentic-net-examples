using System;
using System.Text;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsExamples
{
    public class CustomXmlAndStyleCleanupDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook wb = new Workbook();

                // -------------------------------------------------
                // 1. Add sample data with distinct styles
                // -------------------------------------------------
                Worksheet sheet = wb.Worksheets[0];
                for (int i = 0; i < 5; i++)
                {
                    // Populate a cell
                    Cell cell = sheet.Cells[i, 0];
                    cell.PutValue($"Item {i + 1}");

                    // Create a unique style for each cell
                    Style style = wb.CreateStyle();
                    style.Font.Name = "Arial";
                    style.Font.Size = 10 + i;
                    style.Font.IsBold = (i % 2 == 0);
                    style.Font.Color = (i % 2 == 0) ? System.Drawing.Color.Blue : System.Drawing.Color.Green;

                    // Apply the style
                    cell.SetStyle(style);
                }

                // Delete some rows to make some styles unused
                sheet.Cells.DeleteRows(3, 2); // rows 3 and 4 are removed, their styles become unused

                // -------------------------------------------------
                // 2. Add a custom XML part (data + optional schema)
                // -------------------------------------------------
                string xmlData = "<MyData xmlns=\"http://example.com\"><Item>Value</Item></MyData>";
                string xmlSchema = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                                 + "<xs:schema xmlns:xs=\"http://www.w3.org/2001/XMLSchema\""
                                 + " targetNamespace=\"http://example.com\""
                                 + " xmlns=\"http://example.com\" elementFormDefault=\"qualified\">"
                                 + "<xs:element name=\"MyData\"><xs:complexType><xs:sequence>"
                                 + "<xs:element name=\"Item\" type=\"xs:string\"/>"
                                 + "</xs:sequence></xs:complexType></xs:element>"
                                 + "</xs:schema>";

                // Convert strings to UTF‑8 byte arrays
                byte[] dataBytes = Encoding.UTF8.GetBytes(xmlData);
                byte[] schemaBytes = Encoding.UTF8.GetBytes(xmlSchema);

                // Add the custom XML part to the workbook
                int xmlPartIndex = wb.CustomXmlParts.Add(dataBytes, schemaBytes);

                // Optionally set a custom ID for the part
                CustomXmlPart part = wb.CustomXmlParts[xmlPartIndex];
                part.ID = Guid.NewGuid().ToString();

                // -------------------------------------------------
                // 3. Remove all unused styles in one step
                // -------------------------------------------------
                wb.RemoveUnusedStyles();

                // -------------------------------------------------
                // 4. Save the workbook (the same file contains the custom XML part)
                // -------------------------------------------------
                string outputPath = "CustomXmlAndStyleCleanup.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                wb.Save(outputPath);

                // Inform the user
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
                Console.WriteLine($"Custom XML parts count: {wb.CustomXmlParts.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CustomXmlAndStyleCleanupDemo.Run();
        }
    }
}