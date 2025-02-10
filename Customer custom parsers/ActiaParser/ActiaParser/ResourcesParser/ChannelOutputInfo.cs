using ArtLogics.TestSuite.Boards;
using ArtLogics.TestSuite.Boards.Resources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiaParser.ResourcesParser
{
    public class ChannelOutputInfo
    {
        public Channel Channel { get; set; }
        public string Alias { get; set; }
        public string Function { get; set; }
        public string Resource { get; set; }
        public decimal OffValue { get; set; }
        public string Load { get; set; }
        public object DefaultValue { get; set; }
        public Board Board { get; set; }
        public string Product { get; set; }
        public float OverCurrent { get; set; }
    }
}
