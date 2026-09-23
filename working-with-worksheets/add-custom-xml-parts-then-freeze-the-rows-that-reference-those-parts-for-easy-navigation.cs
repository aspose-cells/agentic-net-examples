// Title: Add custom XML data to an Aspose.Cells worksheet and freeze the header row using C#
// AI Prompts: Write C# code that uses Aspose.Cells to embed a custom XML string, populate cells with employee information, and freeze the first worksheet row. | Show how to parse an XDocument, write its elements to specific cells, and apply FreezePanes on the header in Aspose.Cells for .NET. | Demonstrate creating a new workbook, adding a custom XML part, filling rows from XML, and saving as an .xlsx file with a frozen header.
// Common Searches: asp.net how to import employee XML into Excel with Aspose.Cells and keep header visible | c# Aspose.Cells freeze first row after loading data from XDocument | example of adding custom XML part to workbook and freezing panes in Aspose.Cells | populate worksheet from XML string using Aspose.Cells C# and freeze header row | Aspose.Cells FreezePanes usage after writing XML data to cells
// Tags: custom XML import Aspose.Cells C# | freeze header pane Aspose.Cells FreezePanes | populate worksheet from XDocument Aspose.Cells | save workbook as .xlsx with frozen panes | add custom XML part to Excel using Aspose.Cells

using Aspose.Cells;
using System;
using System.Text;
using System.Xml.Linq;

// The program creates a new workbook, parses a hard‑coded XML string of employee records, writes the ID, Name, and Department values into the first worksheet, freezes the header row, and saves the file as CustomXmlWithFrozenHeader.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Sample custom XML content
            string xmlContent = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Employees>
    <Employee>
        <ID>1</ID>
        <Name>John Doe</Name>
        <Department>Finance</Department>
    </Employee>
    <Employee>
        <ID>2</ID>
        <Name>Jane Smith</Name>
        <Department>HR</Department>
    </Employee>
</Employees>";

            // Write header cells
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Department");

            // Parse XML and import data manually (row index starts at 1 for the second row)
            XDocument doc = XDocument.Parse(xmlContent);
            int currentRow = 1; // zero‑based index; row 1 is the second row in the sheet

            foreach (XElement employee in doc.Root.Elements("Employee"))
            {
                sheet.Cells[currentRow, 0].PutValue((int)employee.Element("ID"));
                sheet.Cells[currentRow, 1].PutValue((string)employee.Element("Name"));
                sheet.Cells[currentRow, 2].PutValue((string)employee.Element("Department"));
                currentRow++;
            }

            // Freeze the first row (header)
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook
            workbook.Save("CustomXmlWithFrozenHeader.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
