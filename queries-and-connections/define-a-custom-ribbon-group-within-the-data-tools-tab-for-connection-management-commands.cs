// Title: Add a custom 'Connection Management' ribbon group with Refresh and Edit buttons to the Data Tools tab using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a Workbook, assigns custom RibbonXml to insert a 'Connection Management' group with large Refresh and Edit buttons into the built‑in TabDataTools, and saves the file as XLSX. | Show how to embed custom ribbon XML in an Aspose.Cells workbook and keep the custom UI after saving.
// Common Searches: aspnet add custom ribbon group to Data Tools tab with Aspose.Cells | c# inject RibbonXml for connection management buttons in Excel workbook | how to preserve custom ribbon UI when saving Aspose.Cells workbook as xlsx | custom UI XML for Excel ribbon using Aspose.Cells .NET example | Aspose.Cells create custom ribbon group for Refresh Connections button
// Tags: set RibbonXml property Aspose.Cells | custom ribbon group Data Tools tab | add connection management buttons Excel ribbon | save workbook with custom UI XLSX | inject custom UI XML Aspose.Cells

using System;
using Aspose.Cells;

// Creates a new Workbook, defines custom Ribbon XML that adds a 'Connection Management' group with large Refresh and Edit buttons to the built‑in Data Tools tab, assigns the XML to Workbook.RibbonXml, and saves the workbook as an XLSX file preserving the custom ribbon UI.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Define custom Ribbon XML.
        // The tab with idMso "TabDataTools" corresponds to the built‑in "Data Tools" tab.
        // Inside it we add a new group "Connection Management" with two command buttons.
        string ribbonXml = @"
<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">
  <ribbon>
    <tabs>
      <tab idMso=""TabDataTools"">
        <group id=""CustomConnGroup"" label=""Connection Management"">
          <button id=""RefreshConn"" label=""Refresh Connections"" size=""large"" onAction=""RefreshConnections"" />
          <button id=""EditConn""   label=""Edit Connections""    size=""large"" onAction=""EditConnections"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        // Assign the XML to the workbook's RibbonXml property (feature rule)
        workbook.RibbonXml = ribbonXml;

        // Save the workbook (lifecycle rule: save)
        // Xlsx format retains the custom ribbon UI.
        workbook.Save("CustomRibbon.xlsx", SaveFormat.Xlsx);
    }
}
