// Title: Add a custom Ribbon tab with a button to an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to embed custom UI XML into an Excel file with Aspose.Cells | How to programmatically add a new ribbon tab in Excel using Aspose.Cells .NET | Aspose.Cells Workbook.CustomXmlParts example for custom ribbon customization | Insert custom ribbon button into .xlsx via Aspose.Cells C# | Save Excel workbook with custom ribbon UI using Aspose.Cells library
// Tags: Aspose.Cells Workbook.CustomXmlParts custom ribbon | C# add custom ribbon tab to Excel workbook | embed custom UI XML in .xlsx with Aspose.Cells | programmatic Excel ribbon customization .NET | custom ribbon button using Aspose.Cells API

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads or creates an Excel workbook, defines a custom UI XML that describes a new ribbon tab containing a button, adds this XML as a custom part via Workbook.CustomXmlParts, and saves the workbook, demonstrating how to embed custom ribbon UI into an .xlsx file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input workbook exists; if not, create a new one.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook
                workbook.Save(inputPath);   // optionally save the placeholder file
            }

            // -------------------------------------------------
            // Customize the Ribbon (add a custom UI XML part)
            // -------------------------------------------------
            // This XML defines a simple custom tab with a button.
            // Adjust the XML as needed for your ribbon customization.
            string customUIXml = @"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
  <ribbon>
    <tabs>
      <tab id='MyCustomTab' label='My Tab'>
        <group id='MyGroup' label='My Group'>
          <button id='MyButton' label='Click Me' size='large' onAction='OnMyButtonClick' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

            // Add the custom UI part to the workbook.
            // Convert both the part name and XML content to byte arrays as required by the API.
            workbook.CustomXmlParts.Add(
                Encoding.UTF8.GetBytes("customUI"),
                Encoding.UTF8.GetBytes(customUIXml));

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
