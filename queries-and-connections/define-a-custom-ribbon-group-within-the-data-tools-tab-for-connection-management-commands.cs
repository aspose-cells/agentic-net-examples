// Title: Add a custom "Connection Management" ribbon group with Refresh and Edit buttons to the Data Tools tab using Aspose.Cells for .NET
// Common Searches: how to add a custom ribbon group to the Data Tools tab in Excel with Aspose.Cells C# | embed custom UI XML in a workbook to create connection management buttons using Aspose.Cells | Aspose.Cells C# add Refresh Connections button to Excel ribbon via custom XML part | create custom ribbon group for data connections in an XLSX file with Aspose.Cells | save custom UI XML as a custom XML part in an Aspose.Cells workbook
// Tags: custom ribbon group Data Tools Aspose.Cells | embed custom UI XML workbook | add connection management buttons Excel ribbon | custom XML part Aspose.Cells C# | refresh connections button custom UI

using Aspose.Cells;
using System;
using System.Text;

// The example creates a new Workbook, defines a Custom UI XML snippet that adds a "Connection Management" group containing "Refresh Connections" and "Edit Connections" buttons to the built‑in Data Tools tab, inserts the XML as a UTF‑8 custom XML part, and saves the workbook as WorkbookWithCustomRibbon.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Custom UI XML that adds a group to the built‑in "Data Tools" tab (TabData).
            string customUiXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab idMso=""TabData"">
        <group id=""ConnectionManagementGroup"" label=""Connection Management"">
          <button id=""RefreshConnections"" 
                  label=""Refresh Connections"" 
                  imageMso=""RefreshAll"" 
                  size=""large"" 
                  onAction=""OnRefreshConnections""/>
          <button id=""EditConnections"" 
                  label=""Edit Connections"" 
                  imageMso=""DataConnectionProperties"" 
                  size=""large"" 
                  onAction=""OnEditConnections""/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

            // Convert the XML string to a UTF‑8 byte array and add it as a custom XML part.
            byte[] customUiBytes = Encoding.UTF8.GetBytes(customUiXml);
            workbook.CustomXmlParts.Add(customUiBytes, null);

            // Save the workbook.
            workbook.Save("WorkbookWithCustomRibbon.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
