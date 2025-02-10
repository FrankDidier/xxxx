using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArtLogics.Translation.Parser.Exception
{
    public enum ParserExceptionType
    {
        RESOURCE_NOT_FOUND,
        VALUE_NOT_FOUND,
        RELAY_LOGIC_NOT_FOUND,
        CALCULATION_ISSUE,
        RESULT_VALUE_EXPECTED_NOT_FOUND,
        CAN_CHANNEL_NOT_FOUND,
        TEST_DEFINITION_NOT_FOUND,
        CAN_LIMIT_DEFINITION_ERROR,
        CHANNEL_NOT_FOUND,
        RESULT_CHANNEL_NOT_FOUND,
        PREREQUIRE_NOT_FOUND,
        BAD_NUMBER_FORMAT,
        LOGIC_NOT_FOUND,
        FORMULAT_LOGIC_NOT_FOUND,
        CAN_CALCULATION_DIVISION_BY_ZERO,
        NULL_REMARK_FOR_SYMBOL_GAUGE,
        WIPER_WASH_ISSUE,
        USERACTION_NOT_WELL_FORMATED,
        STATE_NOT_DEFINED,
        RESISTIVE_VALUE_UNAVAILABLE
    }
}
