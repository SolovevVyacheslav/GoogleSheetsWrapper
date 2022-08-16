using Google.Apis.Sheets.v4.Data;
using System.Linq;

namespace GoogleSheetsWrapper.Utils
{
    /// <summary>
    /// Provides converters and utils for GSheets' fields validation
    /// </summary>
    internal class ValidationUtils
    {
        public static DataValidationRule ConvertToValidationRule(SheetFieldValidationAttribute validationAttribute)
        {
            return new DataValidationRule
            {
                Strict = validationAttribute.Strict,
                InputMessage = validationAttribute.InputMessage,
                ShowCustomUi = validationAttribute.ShowCustomUi,
                Condition = GetCondition(
                    validationAttribute.ConditionType,
                    validationAttribute.UserEnteredValues,
                    validationAttribute.RelativeDate)
            };
        }

        private static BooleanCondition GetCondition(ConditionType conditionType, string[] userEnteredValues, string relativeDate)
        {
            return new BooleanCondition
            {
                Type = conditionType.ToString(),
                Values = userEnteredValues?.Select(userValue => new ConditionValue
                {
                    RelativeDate = relativeDate,
                    UserEnteredValue = userValue
                })
                .ToList()
            };
        }
    }
}
