using System;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaReferenceFromJson
{
    static void Main()
    {
        // JSON that defines the VBA project reference to be added
        string json = @"{
            ""Name"": ""MyProject"",
            ""AbsoluteLibid"": ""C:\\Projects\\MyProject.xlam"",
            ""RelativeLibid"": ""..\\MyProject.xlam""
        }";

        // Parse the JSON and extract the required fields
        using JsonDocument doc = JsonDocument.Parse(json);
        JsonElement root = doc.RootElement;
        string name = root.GetProperty("Name").GetString();
        string absoluteLibid = root.GetProperty("AbsoluteLibid").GetString();
        string relativeLibid = root.GetProperty("RelativeLibid").GetString();

        // Create a new workbook (macro‑enabled) and ensure it has a VBA project
        Workbook workbook = new Workbook();
        workbook.Save("temp.xlsm", SaveFormat.Xlsm);          // creates the VBA project
        workbook = new Workbook("temp.xlsm");                // reload to access the VBA project

        // Add the external VBA project reference using Aspose.Cells.Vba API
        VbaProject vbaProject = workbook.VbaProject;
        vbaProject.References.AddProjectRefrernce(name, absoluteLibid, relativeLibid);

        // Save the workbook with the added reference
        workbook.Save("WorkbookWithReference.xlsm", SaveFormat.Xlsm);
    }
}