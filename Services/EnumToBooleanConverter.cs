using System;
using System.Globalization;
using System.Windows.Data;

namespace ValveFlangeMulti.Services
{
    public sealed class EnumToBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null)
                return false;

            try
            {
                return string.Equals(
                    value.ToString(),
                    parameter.ToString(),
                    StringComparison.OrdinalIgnoreCase
                );
            }
            catch
            {
                // 변환 실패 시 false 반환
                return false;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (parameter == null || targetType == null)
                return Binding.DoNothing;

            try
            {
                return Enum.Parse(targetType, parameter.ToString());
            }
            catch
            {
                // 파싱 실패 시 변환 안함
                return Binding.DoNothing;
            }
        }
    }
}
