using ArtLogics.TestSuite.Limits;
using ArtLogics.TestSuite.Limits.Comparisons;
using ArtLogics.TestSuite.TestResults;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArtLogics.Translation.Parser.Utils
{
    public static class ProjectUtilsLimitFunction
    {
        public static ErrorMessage BuildErrorMessage(string name, Severity severity)
        {
            var errorMessage = new ErrorMessage();
            errorMessage.Name = name;
            errorMessage.Severity = severity;

            return errorMessage;
        }

        public static Limit BuildLimit(ErrorMessage errorMessage, string name, Comparison comparison, ComparisonKind comparisonKind, int id)
        {
            var limit = new Limit();
            limit.ErrorMessage = errorMessage;
            limit.Name = name;
            limit.Container.Comparison = comparison;
            limit.Container.ComparisonKind = comparisonKind;
            limit.Container.Id = id;

            return limit;
        }
    }
}
