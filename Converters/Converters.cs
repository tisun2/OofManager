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

/// <summary>
/// Inverse of <see cref="BoolToVisibilityConverter"/>: false → Visible, true → Collapsed.
/// Lets a single bool drive both "show A" and "show B" without separate properties.
/// </summary>
public class InvertBoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object? parameter, CultureInfo culture)
        => value is bool b && !b ? Visibility.Visible : Visibility.Collapsed;

    public object ConvertBack(object value, Type targetType, object? parameter, CultureInfo culture)
        => value is Visibility v && v != Visibility.Visible;
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
    {
        if (value is not string raw)
            return Binding.DoNothing;

        var s = raw.Trim();
        if (s.Length == 0)
            return Binding.DoNothing;

        // Accept "9", "9:30", "09:30", "9.30", "9 30", trailing am/pm tolerated.
        var lower = s.ToLowerInvariant();
        bool isPm = false, isAm = false;
        if (lower.EndsWith("pm")) { isPm = true; s = s.Substring(0, s.Length - 2).TrimEnd(); }
        else if (lower.EndsWith("am")) { isAm = true; s = s.Substring(0, s.Length - 2).TrimEnd(); }

        var parts = s.Split(new[] { ':', '.', ' ' }, StringSplitOptions.RemoveEmptyEntries);
        int hours, minutes = 0;
        if (parts.Length == 1)
        {
            if (!int.TryParse(parts[0], NumberStyles.Integer, culture, out hours))
                return Binding.DoNothing;
        }
        else if (parts.Length == 2)
        {
            if (!int.TryParse(parts[0], NumberStyles.Integer, culture, out hours) ||
                !int.TryParse(parts[1], NumberStyles.Integer, culture, out minutes))
                return Binding.DoNothing;
        }
        else
        {
            return Binding.DoNothing;
        }

        if (isPm && hours < 12) hours += 12;
        if (isAm && hours == 12) hours = 0;

        if (hours < 0 || hours > 24 || minutes < 0 || minutes > 59)
            return Binding.DoNothing;
        if (hours == 24 && minutes != 0)
            return Binding.DoNothing;

        return new TimeSpan(hours, minutes, 0);
    }
}
