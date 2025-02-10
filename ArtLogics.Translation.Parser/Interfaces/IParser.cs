using ArtLogics.Translation.Parser.Model;
using ArtLogics.Translation.Parser.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArtLogics.Translation.Parser.Interfaces
{
    public interface IParser
    {
        void Parse(string inputFile, string outputFile, ProjectConfiguration projectConfig);
        void Dispose();
    }
}
