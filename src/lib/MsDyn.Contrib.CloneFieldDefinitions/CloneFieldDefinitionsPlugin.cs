using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace MsDyn.Contrib.CloneFieldDefinitions
{
    [Export(typeof(IXrmToolBoxPlugin)),
    ExportMetadata("BackgroundColor", "MediumBlue"),
    ExportMetadata("PrimaryFontColor", "White"),
    ExportMetadata("SecondaryFontColor", "LightGray"),
    ExportMetadata("SmallImageBase64", ""),
    ExportMetadata("BigImageBase64", ""),
    ExportMetadata("Name", "Clone Field Definitions"),
    ExportMetadata("Description", "A plugin for cloning entity field definitions from one entity to another")]
    public class CloneFieldDefinitionsPlugin : PluginBase
    {
        public override IXrmToolBoxPluginControl GetControl()
        {
            return new CloneFieldDefinitionsControl();
        }
    }
}
