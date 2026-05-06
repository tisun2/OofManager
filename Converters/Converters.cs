using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace OofManager.Wpf.Converters;

public class InvertBoolConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object? parameter, CultureInfo culture)
        => value is bool b ? !b : DependencyProperty.UnsetValue;

    public object ConvertBack(object value, Type targetType, object? parameter, CultureInfo culture)
        => value is bool b ? !b : DependencyProperty.UnsetValue;
}

public class BoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object? parameter, CultureInfo culture)
        => value is bool b && b ? Visibility.Visible : Visibility.Collapsed;

    public object ConvertBack(object value, Type targetType, object? parameter, CultureInfo culture)
        => value is Visibility v && v == Visibility.Visible;
}

public class OofStatusToTextConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is Models.OofStatus status)
        {
            return status switch
            {
                Models.OofStatus.Enabled => "Enabled",
                Models.OofStatus.Scheduled => "Scheduled",
                _ => "Disabled"
            };
        }
        return "Disabled";
    }

    public object ConvertBack(object value, Type targetType, object? parameter, CultureInfo culture)
        => Binding.DoNothing;
}

/// <summary>
/// One-way formatter: TimeSpan → "HH:mm" for display in time pickers/dropdowns.
/// </summary>
public class TimeSpanToHourMinuteConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object? parameter, CultureInfo culture)
        => value is TimeSpan ts ? ts.ToString(@"hh\:mm", culture) : string.Empty;

    public object ConvertBack(object value, Type targetType, object? parameter, CultureInfo culture)
        => Binding.DoNothing;
}
