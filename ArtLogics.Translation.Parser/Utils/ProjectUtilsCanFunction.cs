using ArtLogics.TestSuite.Environment.Dbc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArtLogics.Translation.Parser.Utils
{
    public static class ProjectUtilsCanFunction
    {
        public static SignalData BuildSignalData(double lowLimit, double highLimit, double offset, double scale)
        {
            var data = new SignalData();
            data.HighLimit = highLimit;
            data.LowLimit = lowLimit;
            data.Offset = offset;
            data.Scale = scale;

            return data;
        }

        public static SignalBitInfo BuildSignalBitInfo(int length, int position, SignalEndian endian)
        {
            var bitInfo = new SignalBitInfo();
            bitInfo.Length = length;
            bitInfo.Position = position;
            bitInfo.Endianess = endian;

            return bitInfo;
        }
    }
}
