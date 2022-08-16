using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace GoogleSheetsWrapper.Attributes
{
    internal class SheetFieldAttributesComparer : IComparer<AttributesContainer>
    {
        public int Compare([AllowNull] AttributesContainer x, [AllowNull] AttributesContainer y)
        {
            if (x != null && y != null)
            {
                //only ColumnID identifies property as a unique entity.
                return x.SheetField.ColumnID.CompareTo(y.SheetField.ColumnID);
            }

            return 0;
        }
    }
}
