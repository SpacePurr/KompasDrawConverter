using System.Collections.Generic;

namespace KompasDrawConverter
{
    public static class Constants
    {
        public static Dictionary<string, ConverterMode> LocalizationConverterMode => new Dictionary<string, ConverterMode>()
        {
            ["Конвертация в PDF"] = ConverterMode.ToPdf,
            ["Конвертация в DWG"] = ConverterMode.ToDwg
        };
    }
}