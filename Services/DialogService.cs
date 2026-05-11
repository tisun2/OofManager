using System.Windows;
using ModernWpf.Controls;

namespace OofManager.Wpf.Services;

public class DialogService : IDialogService
{
    public async Task<bool> ConfirmAsync(string title, string message, string accept = "OK", string cancel = "Cancel")
    {
        var dlg = new ContentDialog
        {
            Title = title,
            Content = message,
            PrimaryButtonText = accept,
            CloseButtonText = cancel,
            DefaultButton = ContentDialogButton.Primary
        };
        var result = await dlg.ShowAsync();
        return result == ContentDialogResult.Primary;
    }

    public async Task AlertAsync(string title, string message, string close = "OK")
    {
        var dlg = new ContentDialog
        {
            Title = title,
            Content = message,
            CloseButtonText = close
        };
        await dlg.ShowAsync();
    }

    public async Task<DialogChoice> ChoiceAsync(string title, string message, string primary, string secondary, string cancel)
    {
        var dlg = new ContentDialog
        {
            Title = title,
            Content = message,
            PrimaryButtonText = primary,
            SecondaryButtonText = secondary,
            CloseButtonText = cancel,
            DefaultButton = ContentDialogButton.Primary,
        };
        var result = await dlg.ShowAsync();
        return result switch
        {
            ContentDialogResult.Primary => DialogChoice.Primary,
            ContentDialogResult.Secondary => DialogChoice.Secondary,
            _ => DialogChoice.Cancel,
        };
    }

    public async Task<string?> PromptAsync(string title, string message, string accept = "OK", string cancel = "Cancel", string? placeholder = null)
    {
        var stack = new System.Windows.Controls.StackPanel();
        stack.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = message,
            Margin = new Thickness(0, 0, 0, 8),
            TextWrapping = TextWrapping.Wrap
        });
        var input = new System.Windows.Controls.TextBox
        {
            // Keep prompt input bounded so a paste of e.g. 100MB of text
            // can't lock up the UI thread doing layout/measure.
            MaxLength = 256
        };
        if (!string.IsNullOrEmpty(placeholder))
        {
            ModernWpf.Controls.Primitives.ControlHelper.SetPlaceholderText(input, placeholder);
        }
        stack.Children.Add(input);

        var dlg = new ContentDialog
        {
            Title = title,
            Content = stack,
            PrimaryButtonText = accept,
            CloseButtonText = cancel,
            DefaultButton = ContentDialogButton.Primary
        };

        var result = await dlg.ShowAsync();
        return result == ContentDialogResult.Primary ? input.Text : null;
    }
}
