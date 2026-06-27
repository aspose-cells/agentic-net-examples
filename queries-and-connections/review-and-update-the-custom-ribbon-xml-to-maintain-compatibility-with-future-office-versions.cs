using System;
using Aspose.Cells;

class UpdateRibbonXml
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Updated Ribbon XML using the newer 2009/07 namespace for better forward compatibility
        string ribbonXml = 
            @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
                <ribbon>
                    <tabs>
                        <tab id=""customTab"" label=""My Tab"">
                            <group id=""customGroup"" label=""My Group"">
                                <button id=""customButton"" label=""My Button"" size=""large"" onAction=""OnButtonClick"" />
                            </group>
                        </tab>
                    </tabs>
                </ribbon>
            </customUI>";

        // Assign the custom Ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Set OOXML compliance to the strict ISO standard for future Office versions
        workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Save the workbook as a macro‑enabled file (required for custom UI)
        workbook.Save("UpdatedRibbon.xlsm");
    }
}