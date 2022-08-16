using System;

namespace GoogleSheetsWrapper.Attributes
{
    /// <summary>
    /// Field borders attribute
    /// </summary>
    public class SheetFieldBorderAttribute : Attribute
    {
        public SheetFieldBorderAttribute() { }

        /// <summary>
        /// Borders attribute consists of style and color values.
        /// </summary>
        /// <param name="bordersStyle">Positional format: [allBorders] [left, top, right, bottom]  [leftAndRight, topAndBottom]</param>
        /// <param name="rgbaBordersColor">Positional format: [allBorders] [left, top, right, bottom] [leftAndRight, topAndBottom]
        /// Formats: RGBA string 255 255 255 0.9 with space or comma delimeters
        ///          RGB string 255 255 255 with space or comma delimeters (Alpha = 1 by default)
        /// </param>
        public SheetFieldBorderAttribute(BorderStyle[] bordersStyle, string[] rgbaBordersColor)
        {
            BordersStyle = bordersStyle;
            RgbaBordersColor = rgbaBordersColor;
        }

        /// <summary>
        /// Borders attribute consists of style and color values.
        /// </summary>
        /// <param name="bordersStyle">Positional format: [allBorders] [left, top, right, bottom]  [leftAndRight, topAndBottom]</param>
        public SheetFieldBorderAttribute(BorderStyle[] bordersStyle)
        {
            BordersStyle = bordersStyle;
        }

        /// <summary>
        /// [allBorders]
        /// [left, top, right, bottom]
        /// [leftAndRight, topAndBottom]
        /// </summary>
        public BorderStyle[] BordersStyle { get; set; }

        /// <summary>
        /// [allBorders]
        /// [left, top, right, bottom]
        /// [leftAndRight, topAndBottom]
        /// </summary>
        public string[] RgbaBordersColor { get; set; }
    }

    public enum BorderStyle
    {
        //  The style is not specified. Do not use this.
        STYLE_UNSPECIFIED,

        //The border is dotted.
        DOTTED,

        // The border is dashed.
        DASHED,

        //The border is a thin solid line.
        SOLID,

        //The border is a medium solid line.
        SOLID_MEDIUM,

        // The border is a thick solid line.
        SOLID_THICK,

        // No border. Used only when updating a border in order to erase it.
        NONE,

        //The border is two solid lines.
        DOUBLE
    }
}
