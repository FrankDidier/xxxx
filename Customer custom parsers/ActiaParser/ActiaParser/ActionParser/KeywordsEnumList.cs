using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiaParser.ActionParser
{
    public enum KeywordsEnum
    {
        NOKEYWORD = 0,
        FLASH = 1,
        ON = 2,
        OFF = 4,
        DURING = 8,
        AFTER = 16,
        INF = 32,
        EQUAL = 64,
        SUP = 128,
        STOP = 256,
        SECOND = 512,
        MINUTES = 1024,
        MILLISECOND = 2048,
        AVG = 4096,
        HIGH = 8192,
        LOW = 16384,
        WIPER_WASH = 32768,
        WIPER_INTERVAL = 65536,
        OC = 131072,
        GND = 262144,
        HIGHLOW = 524288,
        EVERY = 1048576,
    }

    public static class KeywordsEnumList
    {
        private static Dictionary<string, KeywordsEnum> keywords = new Dictionary<string, KeywordsEnum>()
        {
            { KeywordsEnum.FLASH.ToString(), KeywordsEnum.FLASH},
            { KeywordsEnum.ON.ToString(), KeywordsEnum.ON},
            { KeywordsEnum.OFF.ToString(), KeywordsEnum.OFF},
            { KeywordsEnum.DURING.ToString(), KeywordsEnum.DURING},
            { KeywordsEnum.AFTER.ToString(), KeywordsEnum.AFTER},
            { "<", KeywordsEnum.INF},
            { "=", KeywordsEnum.EQUAL},
            { ">", KeywordsEnum.SUP},
            { KeywordsEnum.STOP.ToString(), KeywordsEnum.STOP},
            { "s", KeywordsEnum.SECOND},
            { "min", KeywordsEnum.MINUTES},
            { "ms", KeywordsEnum.MILLISECOND},
            { KeywordsEnum.AVG.ToString(), KeywordsEnum.AVG},
            { KeywordsEnum.OC.ToString(), KeywordsEnum.OC},
            { KeywordsEnum.GND.ToString(), KeywordsEnum.GND},
            { KeywordsEnum.HIGHLOW.ToString(), KeywordsEnum.HIGHLOW},
        };

        public static Dictionary<KeywordsEnum, string> KeyWordToString = new Dictionary<KeywordsEnum, string>()
        {
            { KeywordsEnum.INF, "<"},
            { KeywordsEnum.EQUAL, "="},
            { KeywordsEnum.SUP, ">"},
            { KeywordsEnum.SECOND, "s"},
            { KeywordsEnum.MINUTES, "min"},
            { KeywordsEnum.MILLISECOND, "ms"},
        };

        public static KeywordsEnum GetKeyword(string test)
        {
            var result = KeywordsEnum.NOKEYWORD;
            foreach (var keyword in keywords)
            {
                if (test.Contains(keyword.Key.ToLower())) {
                    result |= keyword.Value;
                }
            }

            return result;
        }

        public static float CalculateTIme(string test, float number)
        {
            if (test.Substring(test.Count() - 2) == "ms") {
            }
            else if (test[test.Count() - 1] == 's')
            {
                number *= 1000;
            }
            else if (test.Substring(test.Count() - 3) == "min")
            {
                number *= 1000 * 60;
            }

            return number;
        }

    }
}
