using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlCommentDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook wb = new Workbook();

                // Sample XML content
                string xml = @"<Transmittals><Issued_Document>Doc1</Issued_Document><Issued_Document>Doc2</Issued_Document></Transmittals>";

                // Write XML to a temporary file (required by Aspose.Cells XmlMaps.Add)
                string tempXmlPath = Path.Combine(Path.GetTempPath(), "TempXmlMap.xml");
                File.WriteAllText(tempXmlPath, xml);

                // Add the XML map from the temporary file
                int mapIndex = wb.Worksheets.XmlMaps.Add(tempXmlPath);
                XmlMap xmlMap = wb.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "Transmittals_Map";

                // Clean up the temporary file
                if (File.Exists(tempXmlPath))
                {
                    File.Delete(tempXmlPath);
                }

                // Get the first worksheet
                Worksheet sheet = wb.Worksheets[0];
                Cells cells = sheet.Cells;

                // Define cells to link with XML elements
                var linkedCells = new[]
                {
                    new { Row = 0, Column = 0, Path = "/Transmittals/Issued_Document[1]" },
                    new { Row = 1, Column = 0, Path = "/Transmittals/Issued_Document[2]" }
                };

                // Link cells and add visible comments showing the XPath
                foreach (var item in linkedCells)
                {
                    cells.LinkToXmlMap(xmlMap.Name, item.Row, item.Column, item.Path);

                    int commentIdx = sheet.Comments.Add(item.Row, item.Column);
                    Comment comment = sheet.Comments[commentIdx];
                    comment.Note = $"XPath: {item.Path}";
                    comment.IsVisible = true;
                }

                // Save the workbook
                string outputPath = "XmlLinkedComments.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}