using Google.Apis.Sheets.v4.Data;
using GoogleSheetsWrapper.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace GoogleSheetsWrapper.Utils
{
    public class SheetFieldAttributeUtils
    {
        public static void PopulateRecord<T>(T record, IList<object> row, int minColumnId = 1) where T : BaseRecord
        {
            var properties = record.GetType().GetProperties();

            foreach (var property in properties)
            {
                var attribute = property.GetCustomAttribute<SheetFieldAttribute>();

                if (attribute != null && row.Count > attribute.ColumnID - minColumnId)
                {
                    var stringValue = row[attribute.ColumnID - minColumnId]?.ToString();

                    if (attribute.FieldType == SheetFieldType.String)
                    {
                        property.SetValue(record, stringValue);
                    }
                    else if (attribute.FieldType == SheetFieldType.Currency)
                    {
                        if (!string.IsNullOrWhiteSpace(stringValue))
                        {
                            var value = CurrencyParsing.ParseCurrencyString(stringValue);
                            property.SetValue(record, value);
                        }
                    }
                    else if (attribute.FieldType == SheetFieldType.PhoneNumber)
                    {
                        var value = PhoneNumberParsing.RemoveUSInterationalPhoneCode(stringValue);

                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            property.SetValue(record, long.Parse(value));
                        }
                    }
                    else if (attribute.FieldType == SheetFieldType.DateTime)
                    {
                        DateTime? dt = new DateTime();

                        if (!string.IsNullOrWhiteSpace(stringValue))
                        {
                            var serialNumber = double.Parse(stringValue);

                            dt = DateTimeUtils.ConvertFromSerialNumber(serialNumber);

                            property.SetValue(record, dt);
                        }
                    }
                    else if (attribute.FieldType == SheetFieldType.Number)
                    {
                        if (!string.IsNullOrWhiteSpace(stringValue))
                        {
                            var value = double.Parse(stringValue);
                            property.SetValue(record, value);
                        }
                    }
                    else if (attribute.FieldType == SheetFieldType.Integer)
                    {
                        if (!string.IsNullOrWhiteSpace(stringValue))
                        {
                            int value = int.Parse(stringValue);
                            property.SetValue(record, value);
                        }
                    }
                    else if (attribute.FieldType == SheetFieldType.Boolean)
                    {
                        bool boolValue = Convert.ToBoolean(stringValue);
                        property.SetValue(record, boolValue);
                    }
                    else
                    {
                        throw new ArgumentException($"{attribute.FieldType} not supported yet!");
                    }
                }
            }
        }

        public static CellData GetCellDataForSheetField<T>(T record, AttributesContainer attributesContainer, PropertyInfo property)
        {
            var cell = new CellData
            {
                UserEnteredValue = new ExtendedValue()
            };

            var value = property.GetValue(record);

            cell.UserEnteredValue = EvaluateUserEnteredValue(value, attributesContainer.SheetField.FieldType);
            cell.UserEnteredFormat = EvaluateUserEnteredFormat(attributesContainer);

            if (attributesContainer.SheetFieldValidation != null)
                cell.DataValidation = EvaluateDataValidation(attributesContainer.SheetFieldValidation);

            return cell;
        }

        private static NumberFormat EvaluateNumberFormat(SheetFieldAttribute sheetFieldAttr, SheetFieldType fieldType)
        {
            if (fieldType == SheetFieldType.Currency)
            {
                return new NumberFormat()
                {
                    Pattern = sheetFieldAttr.NumberFormatPattern,
                    Type = "CURRENCY"
                };
            }
            else if (fieldType == SheetFieldType.PhoneNumber)
            {
                return new NumberFormat()
                {
                    Pattern = sheetFieldAttr.NumberFormatPattern,
                    Type = "NUMBER"
                };
            }
            else if (fieldType == SheetFieldType.DateTime)
            {
                return new NumberFormat()
                {
                    Pattern = sheetFieldAttr.NumberFormatPattern,
                    Type = "NUMBER"
                };
            }
            else if (fieldType == SheetFieldType.Number)
            {
                return new NumberFormat()
                {
                    Pattern = sheetFieldAttr.NumberFormatPattern,
                    Type = "NUMBER"
                };
            }
            else if (fieldType == SheetFieldType.Integer)
            {
                return new NumberFormat()
                {
                    Pattern = sheetFieldAttr.NumberFormatPattern,
                    Type = "NUMBER"
                };
            }
            else if (fieldType == SheetFieldType.Boolean)
            {
                return new NumberFormat();
            }
            else if (fieldType == SheetFieldType.String)
            {
                return null;
            }
            else
            {
                throw new ArgumentException($"{fieldType} is not supported yet!");
            }
        }

        private static ExtendedValue EvaluateUserEnteredValue(object value, SheetFieldType fieldType)
        {
            var result = new ExtendedValue();

            if (fieldType == SheetFieldType.String)
            {
                if (value == null)
                {
                    result.StringValue = string.Empty;
                }
                else if (string.IsNullOrWhiteSpace(value.ToString()))
                {
                    result.StringValue = string.Empty;
                }
                else
                {
                    result.StringValue = value.ToString();
                }
            }
            else if (fieldType == SheetFieldType.Currency)
            {
                if (value != null)
                {
                    result.NumberValue = double.Parse(value.ToString());
                }
            }
            else if (fieldType == SheetFieldType.PhoneNumber)
            {
                if (value != null)
                {
                    double parsedNumber = double.Parse(value.ToString());

                    if (parsedNumber != 0)
                    {
                        result.NumberValue = parsedNumber;
                    }
                    else
                    {
                        result.NumberValue = null;
                    }
                }
            }
            else if (fieldType == SheetFieldType.DateTime)
            {
                if (value != null)
                {
                    result.NumberValue = DateTimeUtils.ConvertToSerialNumber((DateTime)value);
                }
            }
            else if (fieldType == SheetFieldType.Number)
            {
                if (value != null)
                {
                    result.NumberValue = double.Parse(value.ToString());
                }
            }
            else if (fieldType == SheetFieldType.Integer)
            {
                if (value != null)
                {
                    result.NumberValue = (int)value;
                }
            }
            else if (fieldType == SheetFieldType.Boolean)
            {
                if (value != null)
                {
                    result.BoolValue = (bool)value;
                }
            }
            else
            {
                throw new ArgumentException($"{fieldType} is not supported yet!");
            }

            return result;
        }

        public static int GetColumnId<T>(Expression<Func<T, object>> expression) where T : BaseRecord
        {
            var attribute = GetSheetFieldAttribute(expression);

            return attribute.ColumnID;
        }

        public static SortedDictionary<AttributesContainer, PropertyInfo> GetAllSheetFieldAttributes<T>()
        {
            return GetAllSheetFieldAttributes(typeof(T));
        }

        public static SortedDictionary<AttributesContainer, PropertyInfo> GetAllSheetFieldAttributes(Type type)
        {
            var result = new SortedDictionary<AttributesContainer, PropertyInfo>(new SheetFieldAttributesComparer());

            var properties = type.GetProperties();

            foreach (var property in properties)
            {
                var attributes = property.GetCustomAttributes();

                if (attributes.Any(a => a is SheetFieldAttribute))
                {
                    result.Add(CreateAttributesContainer(attributes), property);
                }
            }

            return result;
        }

        public static SheetFieldAttribute GetSheetFieldAttribute<T>
            (Expression<Func<T, object>> expression)
        {
            MemberExpression memberExpression;

            if (expression.Body is MemberExpression)
            {
                memberExpression = (MemberExpression)expression.Body;
            }
            else if (expression.Body is UnaryExpression unaryExpression)
            {
                memberExpression = (MemberExpression)unaryExpression.Operand;
            }
            else
            {
                throw new ArgumentException();
            }

            var propertyInfo = (PropertyInfo)memberExpression.Member;
            return propertyInfo.GetCustomAttribute<SheetFieldAttribute>();
        }

        private static AttributesContainer CreateAttributesContainer(IEnumerable<Attribute> attributes)
        {
            var container = new AttributesContainer();
            attributes.ToList().ForEach(attr =>
            {
                switch (attr)
                {
                    case SheetFieldAttribute sheetAttr:
                        container.SheetField = sheetAttr;
                        break;

                    case SheetFieldBorderAttribute sheetBorderAttr:
                        container.SheetFieldBorder = sheetBorderAttr;
                        break;

                    case SheetFieldValidationAttribute sheetValidationAttr:
                        container.SheetFieldValidation = sheetValidationAttr;
                        break;

                    default:
                        break;
                };
            });

            return container;
        }

        private static DataValidationRule EvaluateDataValidation(SheetFieldValidationAttribute sheetFieldValidation)
        {
            if (sheetFieldValidation == null)
            {
                return null;
            }

            return ValidationUtils.ConvertToValidationRule(sheetFieldValidation);
        }

        private static CellFormat EvaluateUserEnteredFormat(AttributesContainer attributesContainer)
        {
            var sheetFieldAttr = attributesContainer.SheetField;
            var bordersAttr = attributesContainer.SheetFieldBorder;

            var fieldType = sheetFieldAttr.FieldType;
            var cellFormat = new CellFormat();

            cellFormat.NumberFormat = EvaluateNumberFormat(sheetFieldAttr, fieldType);

            if (bordersAttr != null)
            {
                cellFormat.Borders = BorderUtils.ConvertToBorders(bordersAttr.BordersStyle, bordersAttr.RgbaBordersColor);
            }

            return cellFormat;
        }
    }
}
