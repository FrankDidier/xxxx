using ArtLogics.TestSuite.Boards;
using ArtLogics.TestSuite.Boards.Resources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiaParser.ResourcesParser
{
    public class ChannelInputInfo
    {
        public Channel Channel { get; set; }
        public string Alias { get; set; }
        public string Function { get; set; }
        public string Resource { get; set; }
        public string HwValue { get; set; }
        public string LogicalContact { get; set; }
        public decimal Pulse { get; set; }
        public decimal Ratio { get; set; }
        public object DefaultValue { get; set; }
        public object CurrentValue { get; set; }
        public Board Board { get; set; }
        public object BaseDefaultValue { get; set; } = null;
    }
}
