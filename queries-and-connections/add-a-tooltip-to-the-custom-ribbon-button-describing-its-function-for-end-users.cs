// Title: How to add a screentip tooltip to a custom ribbon button in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add screentip to custom ribbon button in Excel workbook | set tooltip for custom UI ribbon button using Aspose.Cells .NET | embed custom UI XML with button screentip in Excel file via Aspose.Cells | programmatically add tooltip to Excel ribbon button with Aspose.Cells C# | create custom ribbon group with tooltip using Aspose.Cells API
// Tags: Aspose.Cells custom UI XML part | Excel ribbon button tooltip via Aspose.Cells | C# embed custom UI XML in workbook | screentip attribute Aspose.Cells | programmatic ribbon customization .NET

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The example creates a new Workbook, defines a custom UI XML that includes a ribbon button with a 'screentip' attribute serving as a tooltip, adds this XML as a custom XML part to the workbook, and saves the file as WorkbookWithCustomRibbon.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook();

            // Define the custom UI XML.
            // The <button> element includes a 'screentip' attribute which acts as the tooltip.
            string customUiXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">
  <ribbon>
    <tabs>
      <tab idMso=""TabHome"">
        <group id=""MyCustomGroup"" label=""My Custom Group"">
          <button id=""MyCustomButton""
                  label=""My Button""
                  size=""large""
                  onAction=""OnMyCustomButtonClick""
                  screentip=""Executes the custom action for end users"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

            // Convert XML string to UTF-8 byte array.
            byte[] xmlBytes = Encoding.UTF8.GetBytes(customUiXml);

            // Add the custom UI part to the workbook. Schema data is not required, so pass null.
            workbook.CustomXmlParts.Add(xmlBytes, null);

            // Define output file path.
            string outputPath = "WorkbookWithCustomRibbon.xlsx";

            // Ensure the directory for the output file exists.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the custom ribbon button.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
