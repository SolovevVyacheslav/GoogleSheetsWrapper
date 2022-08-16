namespace GoogleSheetsWrapper.Attributes
{
    /// <summary>
    /// Container for wrapper's attributes
    /// </summary>
    public class AttributesContainer
    {
        public SheetFieldAttribute SheetField { get; set; }

        public SheetFieldBorderAttribute? SheetFieldBorder { get; set; }

        public SheetFieldValidationAttribute? SheetFieldValidation { get; set; } 
    }
}
