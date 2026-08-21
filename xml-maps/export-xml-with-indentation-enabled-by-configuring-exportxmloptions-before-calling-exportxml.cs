// Title: Export Workbook to Indented SpreadsheetML (XML) using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to configure SpreadsheetML2003SaveOptions with IsIndentedFormatting=true to save a workbook as a human‑readable, indented XML file.
// Keywords: Aspose.Cells C# export XML | SpreadsheetML2003SaveOptions indentation | IsIndentedFormatting true | pretty printed XML from Excel | save workbook as XML .NET
// Common Searches: Aspose.Cells enable indentation when saving as XML | C# export Excel to indented SpreadsheetML | How to set IsIndentedFormatting in Aspose.Cells | pretty print XML output from Aspose.Cells workbook
// Developer Intent: Generate a readable, indented SpreadsheetML file from a workbook.
// Use Cases: Version‑control friendly XML export of Excel data | Creating documentation‑ready XML reports | Integrating Excel data with systems that require formatted XML
// AI Prompts: Show C# code to export a workbook to indented SpreadsheetML using Aspose.Cells. | Explain the impact of IsIndentedFormatting on the XML structure and how to adjust indentation depth. | Provide an example that includes error handling while saving an indented XML file.

using Aspose.Cells;
using System;

// Demonstrates how to configure SpreadsheetML2003SaveOptions with IsIndentedFormatting=true to save a workbook as a human‑readable, indented XML file.
class ExportXmlWithIndentation
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Employees";

            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("John");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Alice");

            // Save the workbook as SpreadsheetML (XML) with indentation enabled
            SpreadsheetML2003SaveOptions saveOptions = new SpreadsheetML2003SaveOptions
            {
                IsIndentedFormatting = true // Enable child element indentation
            };
            workbook.Save("IndentedOutput.xml", saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
