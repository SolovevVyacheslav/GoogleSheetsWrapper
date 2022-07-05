using Google.Apis.Sheets.v4.Data;
using System;
using System.Linq;

namespace GoogleSheetsWrapper.Utils
{
    /// <summary>
    /// Provides converters and utils for GSheets colors
    /// </summary>
    internal class ColorUtils
    {
        public static ColorStyle ConvertToColorStyle(string stringCodedColor)
        {
            var components = stringCodedColor
                .Split(new char[] { ' ', ',' }, options: StringSplitOptions.RemoveEmptyEntries)
                .Select(c => int.Parse(c))
                .ToList();

            if (components.Count != 3 && components.Count != 4)
            {
                throw new ArgumentException($"Number of color components must be must be either 3 or 4.");
            }

            var isFullRgba = components.Count() == 4;

            return new ColorStyle
            {
                RgbColor = new Color
                {
                    Alpha = isFullRgba ? components[3] : 1,
                    Red = components[0] / 255,
                    Green = components[1] / 255,
                    Blue = components[2] / 255,
                }
            };
        }
    }
}
