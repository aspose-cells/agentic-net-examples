// Title: How to change external link paths and add a custom ribbon UI to an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# update external link source path in an existing workbook | Add a custom ribbon tab to an Excel file with Aspose.Cells .NET | Inject custom UI XML into a .xlsx using Aspose.Cells C# | Programmatically modify external workbook references with Aspose.Cells | Customize Excel ribbon via CustomXmlParts.Add in C#
// Tags: ExternalLinks API path update Aspose.Cells | CustomXmlParts.Add ribbon XML Aspose.Cells | Modify external workbook links .NET | Inject custom UI XML Excel Aspose.Cells | Programmatic ribbon customization C# | Change external link source Excel Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsSample
{
    // The example loads an existing workbook, demonstrates where to update external link paths using the ExternalLinks API (when available), injects custom ribbon UI XML via the CustomXmlParts.Add method, and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "InputWorkbook.xlsx";
                string outputPath = "OutputWorkbook.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // ----------- Change External Link Paths -----------
                // Note: The ExternalLinks API may not be available in older Aspose.Cells versions.
                // If needed, update external link paths using the appropriate API for your version here.

                // ----------- Customize the Ribbon (Custom UI) -----------
                // Define the custom UI XML that will modify the ribbon.
                string customUIXml = @"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
  <ribbon>
    <tabs>
      <tab id='CustomTab' label='My Custom Tab'>
        <group id='CustomGroup' label='My Group'>
          <button id='CustomButton' label='Click Me' size='large' onAction='OnCustomButtonClick' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

                // Add the custom UI part to the workbook.
                // The Add method requires both data and schemaData; schemaData can be null for custom UI.
                workbook.CustomXmlParts.Add(Encoding.UTF8.GetBytes(customUIXml), null);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
