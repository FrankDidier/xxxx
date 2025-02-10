using ArtLogics.TestSuite.Operations;
using ArtLogics.TestSuite.Testing.Actions.Channels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiaParser.ActionParser
{
    public class ChannelActionContainersCompare : IComparer<ChannelActionContainer>
    {
        public int Compare(ChannelActionContainer x, ChannelActionContainer y)
        {
            if (x == null || y == null)
            {
                return 0;
            }

            if (!(x.Operation is RelayOperation) || !(y.Operation is RelayOperation))
            {
                return 0;
            }

            if (((RelayOperation)x.Operation).State == RelayState.CLOSE && ((RelayOperation)y.Operation).State == RelayState.OPEN)
            {
                return 1;
            }

            if (((RelayOperation)x.Operation).State == RelayState.CLOSE && ((RelayOperation)y.Operation).State == RelayState.CLOSE)
            {
                return 0;
            }

            if (((RelayOperation)x.Operation).State == RelayState.OPEN && ((RelayOperation)y.Operation).State == RelayState.OPEN)
            {
                return 0;
            }

            return -1;
        }
    }
}
