using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArtLogics.Translation.Parser.Model
{
    [Serializable]
    public class Extension
    {
        public ExtensionType ExtensionType { get; set; }

        public Extension(ExtensionType ExtensionType)
        {
            this.ExtensionType = ExtensionType;
        }
    }
}
