using Google.Apis.Sheets.v4.Data;
using GoogleSheetsWrapper.Attributes;
using System;
using System.Linq;

namespace GoogleSheetsWrapper.Utils
{
    /// <summary>
    /// Provides converters and utils for GSheets borders
    /// </summary>
    internal class BorderUtils
    {
        public static Borders ConvertToBorders(BorderStyle[] borderStyles, string[] rgbaBordersColors)
        {
            var borders = borderStyles.Length switch
            {
                1 => new Borders
                {
                    Bottom = new Border { Style = borderStyles[0].ToString() },
                    Left = new Border { Style = borderStyles[0].ToString() },
                    Right = new Border { Style = borderStyles[0].ToString() },
                    Top = new Border { Style = borderStyles[0].ToString() },
                },

                2 => new Borders
                {
                    Left = new Border { Style = borderStyles[0].ToString() },
                    Right = new Border { Style = borderStyles[0].ToString() },
                    Bottom = new Border { Style = borderStyles[1].ToString() },
                    Top = new Border { Style = borderStyles[1].ToString() },
                },

                4 => new Borders
                {
                    Left = new Border { Style = borderStyles[0].ToString() },
                    Top = new Border { Style = borderStyles[1].ToString() },
                    Right = new Border { Style = borderStyles[2].ToString() },
                    Bottom = new Border { Style = borderStyles[3].ToString() },
                },

                _ => throw new ArgumentException($"Number of parameteres for {nameof(borderStyles)} must be either 1,2 or 4."),
            };

            if (rgbaBordersColors != null && rgbaBordersColors.Any())
            {
                return borders;
            }

            switch (rgbaBordersColors.Length)
            {
                case 1:
                    var allSidesColorStyle = ColorUtils.ConvertToColorStyle(rgbaBordersColors[0]);
                    borders.Left.ColorStyle = allSidesColorStyle;
                    borders.Top.ColorStyle = allSidesColorStyle;
                    borders.Right.ColorStyle = allSidesColorStyle;
                    borders.Bottom.ColorStyle = allSidesColorStyle;
                    break;

                case 2:
                    var leftRightColosStyle = ColorUtils.ConvertToColorStyle(rgbaBordersColors[0]);
                    var topBottomColorStyle = ColorUtils.ConvertToColorStyle(rgbaBordersColors[1]);
                    borders.Left.ColorStyle = leftRightColosStyle;
                    borders.Right.ColorStyle = leftRightColosStyle;
                    borders.Top.ColorStyle = topBottomColorStyle;
                    borders.Bottom.ColorStyle = topBottomColorStyle;
                    break;

                case 4:
                    borders.Left.ColorStyle = ColorUtils.ConvertToColorStyle(rgbaBordersColors[0]);
                    borders.Top.ColorStyle = ColorUtils.ConvertToColorStyle(rgbaBordersColors[1]);
                    borders.Right.ColorStyle = ColorUtils.ConvertToColorStyle(rgbaBordersColors[2]);
                    borders.Bottom.ColorStyle = ColorUtils.ConvertToColorStyle(rgbaBordersColors[3]);
                    break;

                default:
                    throw new ArgumentException($"Number of parameteres for {nameof(rgbaBordersColors)} must be either 1,2 or 4.");
            };

            return borders;
        }
    }
}
