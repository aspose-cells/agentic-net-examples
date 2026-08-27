// Title: Create a new workbook, enable ISO/IEC 29500:2008 strict OOXML compliance, and save it with the default filename using Aspose.Cells for .NET
// AI Prompts: Generate C# code that instantiates an Aspose.Cells Workbook, sets Settings.Compliance to OoxmlCompliance.Iso29500_2008_Strict, and saves it as StrictComplianceWorkbook.xlsx. | Show how to configure strict OOXML compliance for a newly created workbook before calling Save in Aspose.Cells for .NET. | Adapt existing Aspose.Cells workbook code to enforce ISO 29500 strict mode and use the default file name when saving.
// Common Searches: how to enable ISO 29500 strict compliance in Aspose.Cells C# workbook | Aspose.Cells save workbook with default name after setting OoxmlCompliance | C# example for creating workbook with strict OOXML compliance using Aspose.Cells | set OoxmlCompliance to Iso29500_2008_Strict before saving in Aspose.Cells .NET | default filename saving in Aspose.Cells after configuring strict compliance
// Tags: set OoxmlCompliance Iso29500_2008_Strict Aspose.Cells | save workbook default filename C# | strict OOXML compliance Aspose.Cells | configure workbook compliance .NET | Aspose.Cells ISO 29500 strict mode

using Aspose.Cells;

// Creates a new Aspose.Cells Workbook, configures it for ISO/IEC 29500:2008 strict OOXML compliance, and saves the file as StrictComplianceWorkbook.xlsx using the default filename.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Enable ISO/IEC 29500:2008 Strict OOXML compliance
        workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Save the workbook using a default file name
        workbook.Save("StrictComplianceWorkbook.xlsx");
    }
}
