using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Navigation;
using System.Windows.Threading;
using ModernWpf.Controls;
using OofManager.Wpf.Services;

namespace OofManager.Wpf.Views;

/// <summary>
/// Rich-text editor for an out-of-office reply body, exposing its content as
/// the sanitized HTML fragment Exchange stores.
///
/// The control owns the <see cref="FlowDocument"/>; <see cref="Html"/> is the
/// only thing the view model sees, which keeps <c>MainViewModel</c> free of any
/// WPF document types.
/// </summary>
public partial class HtmlEditor : UserControl
{
    // Fonts the recipient's mail client is overwhelmingly likely to have.
    // Enumerating Fonts.SystemFontFamilies would let a user pick something
    // installed only on their machine, and the reply would silently fall back
    // to a default font for everyone who receives it.
    private static readonly string[] WebSafeFonts =
    {
        "Segoe UI", "Arial", "Calibri", "Times New Roman",
        "Georgia", "Verdana", "Tahoma", "Courier New",
    };

    private static readonly double[] FontSizesPt = { 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32 };

    // Public because the WPF binding engine reflects over the item type from
    // the ComboBox ItemTemplate; a private nested type makes those bindings
    // fail silently at runtime.
    public sealed class PaletteEntry
    {
        public string Name { get; }
        public Brush Brush { get; }
        public Color Color { get; }
        public PaletteEntry(string name, Color color)
        {
            Name = name;
            Color = color;
            Brush = new SolidColorBrush(color);
        }
    }

    private static readonly PaletteEntry[] Palette =
    {
        new("Automatic", Colors.Black),
        new("Dark gray", Color.FromRgb(0x44, 0x44, 0x44)),
        new("Gray", Color.FromRgb(0x80, 0x80, 0x80)),
        new("Red", Color.FromRgb(0xC0, 0x39, 0x2B)),
        new("Orange", Color.FromRgb(0xE6, 0x7E, 0x22)),
        new("Green", Color.FromRgb(0x27, 0x8A, 0x3E)),
        new("Blue", Color.FromRgb(0x00, 0x78, 0xD4)),
        new("Navy", Color.FromRgb(0x1F, 0x37, 0x64)),
        new("Purple", Color.FromRgb(0x7B, 0x2C, 0xBF)),
    };

    // WPF measures font size in device-independent pixels; the toolbar and the
    // generated CSS both talk in points.
    private const double PxPerPt = 96.0 / 72.0;

    // Serializing on every keystroke would run the sanitizer over the whole
    // body for each character typed. A short idle delay collapses a burst of
    // typing into one push without the user noticing a lag.
    private readonly DispatcherTimer _pushTimer = new() { Interval = TimeSpan.FromMilliseconds(250) };

    private bool _suppressPush;   // set while loading the document from Html
    private bool _suppressLoad;   // set while writing Html from the document
    private bool _syncingToolbar; // set while reflecting selection state into the toolbar
    private bool _pastePending;

    public HtmlEditor()
    {
        InitializeComponent();

        FontFamilyBox.ItemsSource = WebSafeFonts;
        FontSizeBox.ItemsSource = FontSizesPt;
        ColorBox.ItemsSource = Palette;

        _pushTimer.Tick += (_, _) =>
        {
            _pushTimer.Stop();
            PushHtml();
        };

        // Pasted content is normalized against the same allowlist the editor
        // emits, so what's on screen is what the recipient will actually get.
        DataObject.AddPastingHandler(Editor, OnPasting);

        // Links inside the editor are chrome, not navigation. Ctrl+click is the
        // built-in gesture; the handler re-validates the scheme because the
        // NavigateUri could have arrived via a paste.
        Editor.AddHandler(Hyperlink.RequestNavigateEvent, new RequestNavigateEventHandler(OnRequestNavigate));

        Editor.PreviewKeyDown += OnEditorPreviewKeyDown;

        Loaded += (_, _) =>
        {
            if (Editor.Document.Blocks.Count == 0)
                FlowDocumentHtml.Load(Editor.Document, Html);
            UpdatePlaceholder();
            SyncToolbarFromSelection();
        };
    }

    // ---------- dependency properties ----------

    public static readonly DependencyProperty HtmlProperty = DependencyProperty.Register(
        nameof(Html), typeof(string), typeof(HtmlEditor),
        new FrameworkPropertyMetadata(
            string.Empty,
            FrameworkPropertyMetadataOptions.BindsTwoWayByDefault,
            OnHtmlPropertyChanged));

    /// <summary>Body-level HTML fragment, sanitized in both directions.</summary>
    public string Html
    {
        get => (string)GetValue(HtmlProperty);
        set => SetValue(HtmlProperty, value);
    }

    public static readonly DependencyProperty PlaceholderProperty = DependencyProperty.Register(
        nameof(Placeholder), typeof(string), typeof(HtmlEditor), new PropertyMetadata(string.Empty));

    public string Placeholder
    {
        get => (string)GetValue(PlaceholderProperty);
        set => SetValue(PlaceholderProperty, value);
    }

    private static void OnHtmlPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        var editor = (HtmlEditor)d;
        if (editor._suppressLoad) return; // echo of our own push

        editor._suppressPush = true;
        try
        {
            FlowDocumentHtml.Load(editor.Editor.Document, e.NewValue as string);
        }
        finally
        {
            editor._suppressPush = false;
        }
        editor.UpdatePlaceholder();
    }

    // ---------- document <-> Html ----------

    private void OnEditorTextChanged(object sender, TextChangedEventArgs e)
    {
        UpdatePlaceholder();
        if (_suppressPush) return;

        if (_pastePending)
        {
            _pastePending = false;
            // Let the paste finish committing before rewriting the document
            // underneath it.
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(NormalizeDocument));
            return;
        }

        _pushTimer.Stop();
        _pushTimer.Start();
    }

    private void PushHtml()
    {
        if (_suppressPush) return;

        var html = FlowDocumentHtml.ToHtml(Editor.Document);
        if (string.Equals(html, Html, StringComparison.Ordinal)) return;

        _suppressLoad = true;
        try
        {
            Html = html;
        }
        finally
        {
            _suppressLoad = false;
        }
    }

    /// <summary>
    /// Round-trips the document through the serializer so anything the editor
    /// cannot represent (tables, images, unsupported styles) disappears
    /// immediately instead of surviving on screen but not in the sent message.
    /// </summary>
    private void NormalizeDocument()
    {
        var html = FlowDocumentHtml.ToHtml(Editor.Document);
        _suppressPush = true;
        try
        {
            FlowDocumentHtml.Load(Editor.Document, html);
            Editor.CaretPosition = Editor.Document.ContentEnd;
        }
        finally
        {
            _suppressPush = false;
        }
        PushHtml();
        UpdatePlaceholder();
    }

    private void OnPasting(object sender, DataObjectPastingEventArgs e) => _pastePending = true;

    private void UpdatePlaceholder()
    {
        var empty = FlowDocumentHtml.IsEmpty(Editor.Document);
        PlaceholderText.Visibility = empty ? Visibility.Visible : Visibility.Collapsed;
    }

    // ---------- toolbar ----------

    private void OnEditorSelectionChanged(object sender, RoutedEventArgs e) => SyncToolbarFromSelection();

    private void SyncToolbarFromSelection()
    {
        _syncingToolbar = true;
        try
        {
            var selection = Editor.Selection;

            BoldButton.IsChecked = selection.GetPropertyValue(TextElement.FontWeightProperty) is FontWeight w
                                   && w >= FontWeights.Bold;
            ItalicButton.IsChecked = selection.GetPropertyValue(TextElement.FontStyleProperty) is FontStyle s
                                     && s != FontStyles.Normal;
            UnderlineButton.IsChecked = selection.GetPropertyValue(Inline.TextDecorationsProperty) is TextDecorationCollection d
                                        && d.Count > 0;

            // A mixed selection returns UnsetValue; leaving the combo blank is
            // the honest representation of "several fonts in here".
            FontFamilyBox.SelectedItem = selection.GetPropertyValue(TextElement.FontFamilyProperty) is FontFamily f
                ? WebSafeFonts.FirstOrDefault(n => string.Equals(n, f.Source, StringComparison.OrdinalIgnoreCase))
                : null;

            FontSizeBox.SelectedItem = selection.GetPropertyValue(TextElement.FontSizeProperty) is double px
                ? FontSizesPt.Cast<double?>().FirstOrDefault(pt => Math.Abs(pt!.Value * PxPerPt - px) < 0.51)
                : null;

            ColorBox.SelectedItem = selection.GetPropertyValue(TextElement.ForegroundProperty) is SolidColorBrush b
                ? Palette.FirstOrDefault(p => p.Color == b.Color)
                : null;
        }
        finally
        {
            _syncingToolbar = false;
        }
    }

    private void ApplyToSelection(DependencyProperty property, object value)
    {
        Editor.Selection.ApplyPropertyValue(property, value);
        Editor.Focus();
        PushHtml();
    }

    private void OnBoldClick(object sender, RoutedEventArgs e)
        => ApplyToSelection(TextElement.FontWeightProperty,
            BoldButton.IsChecked == true ? FontWeights.Bold : FontWeights.Normal);

    private void OnItalicClick(object sender, RoutedEventArgs e)
        => ApplyToSelection(TextElement.FontStyleProperty,
            ItalicButton.IsChecked == true ? FontStyles.Italic : FontStyles.Normal);

    private void OnUnderlineClick(object sender, RoutedEventArgs e)
        => ApplyToSelection(Inline.TextDecorationsProperty,
            UnderlineButton.IsChecked == true ? TextDecorations.Underline : new TextDecorationCollection());

    private void OnFontFamilyChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_syncingToolbar || FontFamilyBox.SelectedItem is not string family) return;
        ApplyToSelection(TextElement.FontFamilyProperty, new FontFamily(family));
    }

    private void OnFontSizeChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_syncingToolbar || FontSizeBox.SelectedItem is not double pt) return;
        ApplyToSelection(TextElement.FontSizeProperty, pt * PxPerPt);
    }

    private void OnColorChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_syncingToolbar || ColorBox.SelectedItem is not PaletteEntry entry) return;
        ApplyToSelection(TextElement.ForegroundProperty, new SolidColorBrush(entry.Color));
    }

    // ---------- hyperlinks ----------

    private void OnEditorPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.K && Keyboard.Modifiers == ModifierKeys.Control)
        {
            e.Handled = true;
            OnInsertLinkClick(sender, e);
        }
    }

    private async void OnInsertLinkClick(object sender, RoutedEventArgs e)
    {
        try
        {
            var selection = Editor.Selection;
            var existing = FindHyperlink(selection.Start) ?? FindHyperlink(selection.End);

            var initialText = existing != null
                ? new TextRange(existing.ContentStart, existing.ContentEnd).Text
                : selection.Text;
            var initialUrl = existing?.NavigateUri?.ToString() ?? string.Empty;

            var result = await PromptForLinkAsync(initialText, initialUrl);
            if (result is null) return;

            var (text, rawUrl) = result.Value;
            var url = ReplyHtml.NormalizeUserUrl(rawUrl);
            if (url is null)
            {
                await new ContentDialog
                {
                    Title = "Can't add that link",
                    Content = "Enter a web address starting with http:// or https://, or an email address.",
                    CloseButtonText = "OK",
                }.ShowAsync();
                return;
            }

            if (string.IsNullOrWhiteSpace(text)) text = url;

            // Editing an existing link is "replace it wholesale": unwrapping
            // first avoids producing a Hyperlink nested inside a Hyperlink,
            // which WPF accepts but cannot render or serialize sensibly.
            var range = existing != null
                ? new TextRange(existing.ElementStart, existing.ElementEnd)
                : new TextRange(selection.Start, selection.End);

            var start = range.Start;
            range.Text = text;
            var written = new TextRange(start, range.End);

            if (Uri.TryCreate(url, UriKind.Absolute, out var uri))
            {
                _ = new Hyperlink(written.Start, written.End)
                {
                    NavigateUri = uri,
                    ToolTip = url,
                };
            }

            Editor.Focus();
            PushHtml();
        }
        catch (Exception ex)
        {
            SyncLogger.Write($"HtmlEditor insert link failed: {ex.GetType().Name}: {ex.Message}");
        }
    }

    private void OnRemoveLinkClick(object sender, RoutedEventArgs e)
    {
        var links = FindHyperlinksIn(Editor.Selection).ToList();
        var atCaret = FindHyperlink(Editor.Selection.Start);
        if (atCaret != null && !links.Contains(atCaret)) links.Add(atCaret);

        foreach (var link in links)
            Unwrap(link);

        Editor.Focus();
        PushHtml();
    }

    private async Task<(string Text, string Url)?> PromptForLinkAsync(string initialText, string initialUrl)
    {
        var textBox = new TextBox { MaxLength = 512, Text = initialText ?? string.Empty };
        ModernWpf.Controls.Primitives.ControlHelper.SetPlaceholderText(textBox, "Text to display");

        var urlBox = new TextBox { MaxLength = 2048, Text = initialUrl ?? string.Empty, Margin = new Thickness(0, 12, 0, 0) };
        ModernWpf.Controls.Primitives.ControlHelper.SetPlaceholderText(urlBox, "https://example.com");

        var panel = new StackPanel();
        panel.Children.Add(new TextBlock { Text = "Text", Margin = new Thickness(0, 0, 0, 4), Opacity = 0.7 });
        panel.Children.Add(textBox);
        panel.Children.Add(new TextBlock { Text = "Address", Margin = new Thickness(0, 12, 0, 4), Opacity = 0.7 });
        panel.Children.Add(urlBox);

        var dialog = new ContentDialog
        {
            Title = string.IsNullOrEmpty(initialUrl) ? "Insert link" : "Edit link",
            Content = panel,
            PrimaryButtonText = "OK",
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Primary,
        };

        var result = await dialog.ShowAsync();
        return result == ContentDialogResult.Primary ? (textBox.Text, urlBox.Text) : null;
    }

    private void OnRequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        e.Handled = true;
        // Never hand the raw NavigateUri to the shell: the document may have
        // been populated from pasted markup, so the scheme is re-checked at the
        // moment of launch rather than trusted from parse time.
        var url = ReplyHtml.SafeUrl(e.Uri?.ToString());
        if (url is null) return;

        try
        {
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            SyncLogger.Write($"HtmlEditor navigate failed: {ex.GetType().Name}: {ex.Message}");
        }
    }

    private static Hyperlink? FindHyperlink(TextPointer? pointer)
    {
        for (var element = pointer?.Parent as DependencyObject; element != null;)
        {
            if (element is Hyperlink link) return link;
            element = element is FrameworkContentElement fce ? fce.Parent : null;
        }
        return null;
    }

    private static IEnumerable<Hyperlink> FindHyperlinksIn(TextRange range)
    {
        var seen = new HashSet<Hyperlink>();
        for (var p = range.Start; p != null && p.CompareTo(range.End) <= 0; p = p.GetNextContextPosition(LogicalDirection.Forward))
        {
            if (p.GetAdjacentElement(LogicalDirection.Forward) is Hyperlink link && seen.Add(link))
                yield return link;
        }
    }

    /// <summary>Replaces a hyperlink with its own content, keeping the text.</summary>
    private static void Unwrap(Hyperlink link)
    {
        var span = new Span();
        foreach (var inline in link.Inlines.ToList())
        {
            link.Inlines.Remove(inline);
            span.Inlines.Add(inline);
        }

        switch (link.Parent)
        {
            case Paragraph paragraph:
                paragraph.Inlines.InsertAfter(link, span);
                paragraph.Inlines.Remove(link);
                break;
            case Span parent:
                parent.Inlines.InsertAfter(link, span);
                parent.Inlines.Remove(link);
                break;
        }
    }
}
