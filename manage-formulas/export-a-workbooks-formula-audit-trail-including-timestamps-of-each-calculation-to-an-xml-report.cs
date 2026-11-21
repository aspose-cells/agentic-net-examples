// Title: Generate an XML report of a workbook’s formula audit trail with calculation timestamps using Aspose.Cells for .NET
// Common Searches: how to export formula audit trail with timestamps from an Excel file using Aspose.Cells C# | Aspose.Cells create XML log of all formula calculations in a workbook | C# generate XML report of formula cells including calculation time with Aspose.Cells | record formula evaluation timestamps in XML using Aspose.Cells .NET | export Excel formula audit data to XML report programmatically
// Tags: Aspose.Cells XML audit log for formulas | C# formula calculation timestamp extraction | generate Excel formula audit XML using Aspose | record formula evaluation metadata Aspose.Cells | export workbook formula audit data .NET

using System;
using System.IO;
using System.Xml;
using Aspose.Cells;

namespace FormulaAuditTrailDemo
{
    // The example loads an Excel workbook, forces a full recalculation, iterates through each worksheet to find formula cells, and builds an XML document that captures the sheet name, cell address, formula text, calculated result, and a UTC timestamp for every calculation before saving the audit trail to a file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputXml = "FormulaAuditTrail.xml";

            try
            {
                // Ensure the input workbook exists; create a simple one if missing.
                if (!File.Exists(inputPath))
                {
                    CreateSampleWorkbook(inputPath);
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Recalculate all formulas in the workbook.
                workbook.CalculateFormula();

                // Build XML audit trail for all formula cells.
                XmlDocument xmlDoc = new XmlDocument();
                XmlElement root = xmlDoc.CreateElement("FormulaAuditTrail");
                xmlDoc.AppendChild(root);

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the used range.
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            if (cell.IsFormula)
                            {
                                XmlElement calcNode = xmlDoc.CreateElement("Calculation");
                                calcNode.SetAttribute("Sheet", sheet.Name);
                                calcNode.SetAttribute("Cell", cell.Name);
                                calcNode.SetAttribute("Formula", cell.Formula);
                                calcNode.SetAttribute("Result", cell.Value?.ToString() ?? string.Empty);
                                // Use the current time as the calculation timestamp.
                                calcNode.SetAttribute("Timestamp", DateTime.UtcNow.ToString("o"));
                                root.AppendChild(calcNode);
                            }
                        }
                    }
                }

                // Save the XML report.
                xmlDoc.Save(outputXml);
                Console.WriteLine($"Audit trail saved to '{outputXml}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Creates a minimal workbook with a sample formula for demonstration purposes.
        private static void CreateSampleWorkbook(string path)
        {
            try
            {
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                ws.Cells["A1"].PutValue(10);
                ws.Cells["A2"].PutValue(20);
                ws.Cells["A3"].Formula = "=A1+A2";
                wb.Save(path);
                Console.WriteLine($"Sample workbook created at '{path}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create sample workbook: {ex.Message}");
                throw;
            }
        }
    }
}
