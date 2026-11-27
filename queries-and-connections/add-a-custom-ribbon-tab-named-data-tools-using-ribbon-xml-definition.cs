// Title: Create a custom "Data Tools" ribbon tab in an Excel macro-enabled (XLSM) workbook using Aspose.Cells for .NET RibbonXml
// Common Searches: how to add a custom ribbon tab to an XLSM file with Aspose.Cells C# | Aspose.Cells RibbonXml example for creating a Data Tools tab in Excel | assign custom UI XML to a workbook and save as macro-enabled using Aspose.Cells | C# code to embed custom ribbon buttons (Import XML, Export CSV) in an Excel workbook with Aspose.Cells | using Aspose.Cells to define custom ribbon groups and save as .xlsm
// Tags: Aspose.Cells RibbonXml custom UI definition | embed custom Excel ribbon in XLSM using .NET | create custom ribbon group with buttons Aspose.Cells | save macro-enabled workbook with custom UI | define custom UI XML for Excel workbook

using System;
using Aspose.Cells;

// The example creates a new workbook, adds sample data, defines Ribbon XML that introduces a "Data Tools" tab containing large Import XML and Export CSV buttons, assigns this XML via the Workbook.RibbonXml property, and saves the file as a macro-enabled XLSM to retain the custom ribbon UI.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Create a new workbook (empty)
            Workbook workbook = new Workbook();

            // Add some sample data to demonstrate that the workbook is functional
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleData";
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alpha");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Beta");

            // Define Ribbon XML that adds a custom tab named "Data Tools"
            // The XML follows the Office 2007+ custom UI schema.
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customDataToolsTab\" label=\"Data Tools\">" +
                "        <group id=\"customDataGroup\" label=\"Data Operations\">" +
                "          <button id=\"btnImportXml\" label=\"Import XML\" size=\"large\" />" +
                "          <button id=\"btnExportCsv\" label=\"Export CSV\" size=\"large\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // Save the workbook as a macro-enabled file (XLSM) so the ribbon XML is retained
            string outputPath = "DataToolsWorkbook.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook successfully saved to '{outputPath}' with a custom ribbon tab \"Data Tools\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
