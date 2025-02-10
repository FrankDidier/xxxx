using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArtLogics.Translation.Parser.Utils
{
    public static class StringUtils
    {
        public static string Reverse(string value) {
            string returnValue = "";

            for (var i = value.Length - 1; i >= 0; i--) {
                returnValue += value[i];
            }

            return returnValue;
        }
    }
}
