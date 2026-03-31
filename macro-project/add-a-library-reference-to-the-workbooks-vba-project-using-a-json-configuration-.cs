using System;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // JSON payload describing the VBA library reference
        string json = @"{
            ""ReferenceType"": ""Registered"",
            ""Name"": ""stdole"",
            ""Libid"": ""*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation"",
            ""AbsoluteLibid"": null,
            ""RelativeLibid"": null,
            ""TwiddledLibid"": null,
            ""ExtendedLibid"": null
        }";

        // Parse the JSON payload
        using JsonDocument doc = JsonDocument.Parse(json);
        JsonElement root = doc.RootElement;

        string referenceType = root.GetProperty("ReferenceType").GetString();
        string name = root.GetProperty("Name").GetString();

        // Create a new workbook (macro‑enabled) and ensure a VBA project exists
        Workbook workbook = new Workbook();
        workbook.Save("temp.xlsm", SaveFormat.Xlsm);          // create .xlsm file
        workbook = new Workbook("temp.xlsm");                // reload to get VBA project

        VbaProject vbaProject = workbook.VbaProject;

        // Add the appropriate reference based on the type specified in JSON
        if (referenceType.Equals("Registered", StringComparison.OrdinalIgnoreCase))
        {
            string libid = root.GetProperty("Libid").GetString();
            vbaProject.References.AddRegisteredReference(name, libid);
        }
        else if (referenceType.Equals("Control", StringComparison.OrdinalIgnoreCase))
        {
            string libid = root.GetProperty("Libid").GetString();
            string twiddledLibid = root.GetProperty("TwiddledLibid").GetString();
            string extendedLibid = root.GetProperty("ExtendedLibid").GetString();
            vbaProject.References.AddControlRefrernce(name, libid, twiddledLibid, extendedLibid);
        }
        else if (referenceType.Equals("Project", StringComparison.OrdinalIgnoreCase))
        {
            string absoluteLibid = root.GetProperty("AbsoluteLibid").GetString();
            string relativeLibid = root.GetProperty("RelativeLibid").GetString();
            vbaProject.References.AddProjectRefrernce(name, absoluteLibid, relativeLibid);
        }

        // Save the workbook with the added VBA reference
        workbook.Save("WorkbookWithReference.xlsm", SaveFormat.Xlsm);
    }
}