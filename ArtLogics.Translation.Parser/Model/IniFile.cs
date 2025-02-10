using ArtLogics.TestSuite.TestResults;
using IniParser;
using IniParser.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArtLogics.Translation.Parser.Model
{
    public class IniFile
    {
        public static IniData IniDataRaw { get; set; }

        public BindingList<Tuple<string, object>> IniData { get; set; }

        public IniFile()
        {
            IniData = new BindingList<Tuple<string, dynamic>>();

            OpenIniFIle("config.ini");
        }

        public void OpenIniFIle(string file)
        {
            var parser = new FileIniDataParser();
            IniDataRaw = parser.ReadFile(file);

            foreach (var section in IniDataRaw.Sections)
            {
                foreach (var key in section.Keys)
                {
                    AnalyzeData(key);
                }
            }
        }

        private void AnalyzeData(KeyData key)
        {
            if (key.KeyName.Contains("VALUE")) {
                IniData.Add(new Tuple<string, dynamic>(key.KeyName, key.Value));
            }
            else if (key.KeyName.Contains("TYPE"))
            {
                IniData.Add(new Tuple<string, dynamic>(key.KeyName, (Severity)Enum.Parse(typeof(Severity), key.Value, true)));
            } else
            {
                IniData.Add(new Tuple<string, dynamic>(key.KeyName, key.Value));
            }
        }
    }
}
