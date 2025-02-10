using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArtLogics.Translation.Parser.Model
{
    [Serializable]
    public class ProjectConfiguration
    {
        public string ProjectName { get; set; }
        public string ProductName { get; set; }

        public BindingList<Resources> Resources { get; set; }

        public ProjectConfiguration()
        {
            ProjectName = IniFile.IniDataRaw["Project"]["Name"];
            ProductName = IniFile.IniDataRaw["Product"]["Name"];

            Resources = new BindingList<Resources>();
            GenerateRessourcesFromIniFile();
        }

        public void GenerateRessourcesFromIniFile()
        {
            for (var i = 1; IniFile.IniDataRaw["Ressource" + i].Count > 0; i++)
            {
                var resourceConfig = IniFile.IniDataRaw["Ressource" + i];
                var type = (ResourceType)Enum.Parse(typeof(ResourceType), resourceConfig["Type"]);
                var resource = new Resources("Ressource" + i, type);

                for (var e = 1; e < 8; e++)
                {
                    ExtensionType ExtensionType = ExtensionType.NONE;
                    if (resourceConfig["Ext" + e] != null && resourceConfig["Ext" + e] != "")
                    {
                        ExtensionType = (ExtensionType)Enum.Parse(typeof(ExtensionType), resourceConfig["Ext" + e].Replace("-", "_"));
                    }
                    var ext = new Extension(ExtensionType);
                    resource.Extensions.Add(ext);
                }

                for (var e = 1; e < 3; e++)
                {
                    if (resourceConfig["CAN" + e] != null)
                    {
                        resource.CanMap.Add("CAN" + e, resourceConfig["CAN" + e]);
                    }
                }
                Resources.Add(resource);
            }
        }
    }
}
