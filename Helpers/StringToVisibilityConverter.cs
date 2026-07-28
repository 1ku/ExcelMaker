using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace ExcelMaker.Helpers;

/// <summary>
/// 字符串为空则折叠，否则显示。
/// </summary>
public class StringToVisibilityConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => string.IsNullOrEmpty(value as string) ? Visibility.Collapsed : Visibility.Visible;

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => Binding.DoNothing;
}
