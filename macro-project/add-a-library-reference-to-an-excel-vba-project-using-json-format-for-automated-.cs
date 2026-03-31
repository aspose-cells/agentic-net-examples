using System;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // JSON describing the VBA project reference to be added
        string json = @"{
            ""Name"": ""MyAddIn"",
            ""AbsoluteLibid"": ""C:\\Addins\\MyAddIn.xlam"",
            ""RelativeLibid"": ""..\\Addins\\MyAddIn.xlam""
        }";

        // Deserialize JSON into a strongly‑typed object
        var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        ReferenceInfo refInfo = JsonSerializer.Deserialize<ReferenceInfo>(json, options);

        // Create a new workbook (initially without VBA)
        Workbook workbook = new Workbook();

        // Save as a macro‑enabled workbook to ensure a VBA project exists,
        // then reload the file so that workbook.VbaProject is initialized.
        workbook.Save("temp.xlsm", SaveFormat.Xlsm);
        workbook = new Workbook("temp.xlsm");

        // Add the external VBA project reference using the parsed values
        int addedIndex = workbook.VbaProject.References.AddProjectRefrernce(
            refInfo.Name,
            refInfo.AbsoluteLibid,
            refInfo.RelativeLibid);

        Console.WriteLine($"Reference added at index {addedIndex}. Total references: {workbook.VbaProject.References.Count}");

        // Save the final workbook with the new reference
        workbook.Save("WorkbookWithReference.xlsm", SaveFormat.Xlsm);
    }

    // Helper class that matches the JSON structure
    private class ReferenceInfo
    {
        public string Name { get; set; }
        public string AbsoluteLibid { get; set; }
        public string RelativeLibid { get; set; }
    }
}