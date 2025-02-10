using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiaParser.ResourcesParser
{
    public class ChannelInputInterpretations : List<ChannelInputInterpretation>
    {
        public decimal GetClosestInferiorValue(decimal x)
        {
            decimal value = 0;
            foreach (var datas in this)
            {
                if (datas.RealValue > 0 && datas.RealValue < x)
                {
                    value = datas.InterpretedValue;
                }
            }

            return value;
        }

        public decimal GetClosestSuperiorValue(decimal x)
        {
            decimal value = 0;

            foreach (var datas in this)
            {
                if (datas.RealValue > 0 && datas.RealValue >= x)
                {
                    value = datas.InterpretedValue;
                }
            }

            return value;
        }

        public decimal Interpollation(decimal x, bool reverse = false) {

            decimal y = -1;
            decimal y0 = -1;
            decimal y1 = -1;
            decimal x0 = -1;
            decimal x1 = -1;

            if (reverse)
            {
                var dataListSort = this.OrderBy(v => v.RealValue).ToList();

                foreach (var datas in dataListSort)
                {
                    if (datas.RealValue > 0 && datas.RealValue >= x)
                    {
                        y1 = datas.InterpretedValue;
                        x1 = datas.RealValue;
                        if (x0 == -1 && y0 == -1)
                        {
                            x0 = dataListSort.Where(v => v.InterpretedValue >= 0).ElementAt(1).InterpretedValue;
                            y0 = dataListSort.Where(v => v.RealValue >= 0).ElementAt(1).RealValue;
                        }
                        break;
                    }

                    if (datas.RealValue > 0 && datas.RealValue <= x)
                    {
                        x1 = x0;
                        y1 = y0;
                        y0 = datas.InterpretedValue;
                        x0 = datas.RealValue;
                    }
                }
            }
            else
            {
                var dataListSort = this.OrderBy(v => v.InterpretedValue).ToList();

                foreach (var datas in dataListSort)
                {
                    if (datas.InterpretedValue >= 0 && datas.InterpretedValue >= x)
                    {
                        x1 = datas.InterpretedValue;
                        y1 = datas.RealValue;
                        if (x0 == -1 && y0 == -1)
                        {
                            x0 = dataListSort.Where(v => v.InterpretedValue >= 0).ElementAt(1).InterpretedValue;
                            y0 = dataListSort.Where(v => v.RealValue >= 0).ElementAt(1).RealValue;
                        }
                        break;
                    }

                    if (datas.InterpretedValue >= 0 && datas.InterpretedValue <= x)
                    {
                        x1 = x0;
                        y1 = y0;
                        x0 = datas.InterpretedValue;
                        y0 = datas.RealValue;
                    }
                }
            }

            var y2 = y1 - y0;
            var x2 = x1 - x0;
            var x3 = x - x0;

            if (x2 == 0)
            {
                return 0;
            }

            var z = y2 / x2;


            y = y0 + x3 * z;

            return y;
        }

        public ChannelInputInterpretation GetCloseEst(decimal value)
        {
            decimal diffUp = decimal.MaxValue;
            decimal diffDown = decimal.MaxValue;
            ChannelInputInterpretation min = new ChannelInputInterpretation();

            foreach (var datas in this)
            {
                if (datas.InterpretedValue > value)
                {
                    diffUp = Math.Abs(value - datas.InterpretedValue);

                    if (diffDown > diffUp)
                    {
                        return min;
                    } else
                    {
                        return datas;
                    }
                }

                if (datas.InterpretedValue <= value)
                {
                    var newDiffDown = Math.Abs(value - datas.InterpretedValue);
                    if (newDiffDown <= diffDown)
                    {
                        min = datas;
                        diffDown = newDiffDown;
                    }
                }
            }

            return min;
        }
    }
}
