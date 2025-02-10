using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArtLogics.Translation.Parser.Model
{
    [Serializable]
    public class Resources
    {
        public string Alias { get; set; }
        public ResourceType RessourceType { get; set; }
        public BindingList<Extension> Extensions { get; set; }
        public Dictionary<string, string> CanMap { get; set; }

        public Resources(String Alias, ResourceType type)
        {
            this.Alias = Alias;
            RessourceType = ResourceType.TCU100;
            Extensions = new BindingList<Extension>();
            CanMap = new Dictionary<string, string>();
        }
    }
}
