using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using OofManager.Wpf.Services;
using OofManager.Wpf.ViewModels;

namespace OofManager.Wpf.Views;

public partial class MainPage : Page
{
    private readonly MainViewModel _vm;

    public MainPage(MainViewModel viewModel)
    {
        InitializeComponent();
        _vm = viewModel;
        DataContext = viewModel;
        Loaded += OnPageLoaded;

        // The DatePicker's inner DatePickerTextBox does not honor the parent's
        // HorizontalContentAlignment, and overriding it via an implicit Style
        // strips the ModernWpf theme. Reach into the templated tree once it's
        // applied and tweak alignment + padding directly to keep the theme.
        VacationStartDatePicker.Loaded += (_, _) => CenterDatePickerText(VacationStartDatePicker);
        VacationEndDatePicker.Loaded += (_, _) => CenterDatePickerText(VacationEndDatePicker);
    }

    // Wraps MainViewModel.LoadAsync so an exception inside it does not bubble
    // out of an `async void` lambda (which would crash the process via the
    // unobserved-exception path). LoadAsync already sets StatusMessage on
    // failure; this guard is purely defensive.
    private async void OnPageLoaded(object sender, RoutedEventArgs e)
    {
        try
        {
            await _vm.LoadAsync();
        }
        catch (Exception ex)
        {
            SyncLogger.Write($"MainPage.OnPageLoaded LoadAsync failed: {ex.GetType().Name}: {ex.Message}");
        }
    }

    private static void CenterDatePickerText(DatePicker picker)
    {
        var textBox = FindDescendant<DatePickerTextBox>(picker);
        if (textBox is null)
            return;

        textBox.HorizontalContentAlignment = HorizontalAlignment.Center;
        textBox.VerticalContentAlignment = VerticalAlignment.Center;
        // Left padding offsets the right-side calendar glyph so the text
        // centers in the visible area; small bottom padding nudges it up to
        // counter the theme's slight top bias.
        textBox.Padding = new Thickness(6, 0, 0, 3);
    }

    private static T? FindDescendant<T>(DependencyObject root) where T : DependencyObject
    {
        var count = VisualTreeHelper.GetChildrenCount(root);
        for (var i = 0; i < count; i++)
        {
            var child = VisualTreeHelper.GetChild(root, i);
            if (child is T match)
                return match;
            var nested = FindDescendant<T>(child);
            if (nested is not null)
                return nested;
        }
        return null;
    }
}
