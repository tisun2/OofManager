using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace OofManager.Wpf.Services;

/// <summary>
/// Converts between the <see cref="FlowDocument"/> a WPF <c>RichTextBox</c>
/// edits and the HTML subset <see cref="ReplyHtml"/> allows on the wire.
///
/// The accepted element set is deliberately identical to
/// <see cref="ReplyHtml.Sanitize"/>'s allowlist. <see cref="Load"/> always runs
/// its input through the sanitizer first, which additionally guarantees
/// well-formed nesting — that is what lets the parser below get away with a
/// simple stack rather than full error recovery.
/// </summary>
public static class FlowDocumentHtml
{
    // WPF sizes text in device-independent pixels (1/96in); HTML/Outlook think
    // in points (1/72in). Replies are authored here and rendered in a mail
    // client, so the conversion has to happen at the boundary or every message
    // goes out a third larger than it looked in the editor.
    private const double PxPerPt = 96.0 / 72.0;

    // ---------- FlowDocument -> HTML ----------

    /// <summary>
    /// Fast check for whether the document contains any non-whitespace text,
    /// stopping at the very first non-whitespace character without allocating
    /// a <see cref="TextRange"/> or full-document string on every keystroke.
    /// </summary>
    public static bool IsEmpty(FlowDocument? document)
    {
        if (document is null) return true;
        return AreBlocksEmpty(document.Blocks);
    }

    private static bool AreBlocksEmpty(BlockCollection blocks)
    {
        foreach (var block in blocks)
        {
            switch (block)
            {
                case Paragraph p:
                    if (!AreInlinesEmpty(p.Inlines)) return false;
                    break;
                case List list:
                    foreach (var item in list.ListItems)
                    {
                        if (!AreBlocksEmpty(item.Blocks)) return false;
                    }
                    break;
                case Section s:
                    if (!AreBlocksEmpty(s.Blocks)) return false;
                    break;
            }
        }
        return true;
    }

    private static bool AreInlinesEmpty(InlineCollection inlines)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case Run run:
                    var text = run.Text;
                    for (var i = 0; i < text.Length; i++)
                    {
                        if (!char.IsWhiteSpace(text[i])) return false;
                    }
                    break;
                case Span span:
                    if (!AreInlinesEmpty(span.Inlines)) return false;
                    break;
            }
        }
        return true;
    }

    /// <summary>
    /// Serializes the document to a sanitized body fragment.
    /// </summary>
    public static string ToHtml(FlowDocument? document)
    {
        if (IsEmpty(document)) return string.Empty;

        var sb = new StringBuilder();
        WriteBlocks(sb, document!.Blocks);
        // Round-trip through the sanitizer so serializer output and wire format
        // can never diverge, and so anything the user pasted that survived into
        // the document still can't escape into the message.
        return ReplyHtml.Sanitize(sb.ToString());
    }

    private static void WriteBlocks(StringBuilder sb, BlockCollection blocks)
    {
        foreach (var block in blocks)
        {
            switch (block)
            {
                case Paragraph p:
                    sb.Append("<p>");
                    var before = sb.Length;
                    WriteInlines(sb, p.Inlines, new InlineFormat(p));
                    // An empty paragraph is the user's blank line. Without a
                    // <br> the browser/mail client collapses it to nothing.
                    if (sb.Length == before) sb.Append("<br>");
                    sb.Append("</p>");
                    break;

                case List list:
                    var tag = list.MarkerStyle is TextMarkerStyle.Decimal
                        or TextMarkerStyle.LowerLatin or TextMarkerStyle.UpperLatin
                        or TextMarkerStyle.LowerRoman or TextMarkerStyle.UpperRoman
                        ? "ol" : "ul";
                    sb.Append('<').Append(tag).Append('>');
                    foreach (var item in list.ListItems)
                    {
                        sb.Append("<li>");
                        WriteBlocks(sb, item.Blocks);
                        sb.Append("</li>");
                    }
                    sb.Append("</").Append(tag).Append('>');
                    break;

                case Section s:
                    WriteBlocks(sb, s.Blocks);
                    break;
            }
        }
    }

    private static void WriteInlines(StringBuilder sb, InlineCollection inlines, InlineFormat inherited)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case LineBreak:
                    sb.Append("<br>");
                    break;

                case Run run when run.Text.Length > 0:
                {
                    var fmt = new InlineFormat(run);
                    var close = OpenFormat(sb, fmt, inherited);
                    sb.Append(WebUtility.HtmlEncode(run.Text));
                    sb.Append(close);
                    break;
                }

                case Hyperlink link:
                {
                    var href = ReplyHtml.SafeUrl(link.NavigateUri?.ToString());
                    var fmt = new InlineFormat(link);
                    // A link's own colour/underline is chrome the mail client
                    // supplies; emitting it would bake the editor's theme into
                    // the recipient's message.
                    var styled = fmt.WithoutLinkChrome();
                    var close = OpenFormat(sb, styled, inherited);
                    if (href != null) sb.Append("<a href=\"").Append(WebUtility.HtmlEncode(href)).Append("\">");
                    WriteInlines(sb, link.Inlines, styled.MergedOver(inherited));
                    if (href != null) sb.Append("</a>");
                    sb.Append(close);
                    break;
                }

                case Span span:
                {
                    var fmt = new InlineFormat(span);
                    var close = OpenFormat(sb, fmt, inherited);
                    WriteInlines(sb, span.Inlines, fmt.MergedOver(inherited));
                    sb.Append(close);
                    break;
                }
            }
        }
    }

    /// <summary>
    /// Emits opening tags for whatever <paramref name="fmt"/> changes relative
    /// to <paramref name="inherited"/> and returns the matching closing tags.
    /// Diffing against the inherited state keeps the markup from accumulating a
    /// redundant style attribute on every single run.
    /// </summary>
    private static string OpenFormat(StringBuilder sb, InlineFormat fmt, InlineFormat inherited)
    {
        var close = new StringBuilder();

        if (fmt.Bold == true && inherited.Bold != true) { sb.Append("<b>"); close.Insert(0, "</b>"); }
        if (fmt.Italic == true && inherited.Italic != true) { sb.Append("<i>"); close.Insert(0, "</i>"); }
        if (fmt.Underline == true && inherited.Underline != true) { sb.Append("<u>"); close.Insert(0, "</u>"); }

        var css = new StringBuilder();
        if (fmt.FontFamily != null && fmt.FontFamily != inherited.FontFamily)
            css.Append("font-family:").Append(fmt.FontFamily).Append(';');
        if (fmt.FontSizePx is double px && Math.Abs(px - (inherited.FontSizePx ?? -1)) > 0.01)
            css.Append("font-size:").Append((px / PxPerPt).ToString("0.#", CultureInfo.InvariantCulture)).Append("pt;");
        if (fmt.Color != null && fmt.Color != inherited.Color)
            css.Append("color:").Append(fmt.Color).Append(';');

        if (css.Length > 0)
        {
            sb.Append("<span style=\"").Append(css).Append("\">");
            close.Insert(0, "</span>");
        }

        return close.ToString();
    }

    /// <summary>
    /// The character-formatting properties this converter round-trips, read
    /// from a <see cref="TextElement"/>'s <i>local</i> values only. Local values
    /// are exactly what <c>TextSelection.ApplyPropertyValue</c> sets, so
    /// ignoring inherited ones keeps the document's implicit defaults out of the
    /// generated markup.
    /// </summary>
    private readonly struct InlineFormat
    {
        public bool? Bold { get; }
        public bool? Italic { get; }
        public bool? Underline { get; }
        public string? FontFamily { get; }
        public double? FontSizePx { get; }
        public string? Color { get; }

        private InlineFormat(bool? bold, bool? italic, bool? underline, string? family, double? sizePx, string? color)
        {
            Bold = bold; Italic = italic; Underline = underline;
            FontFamily = family; FontSizePx = sizePx; Color = color;
        }

        public InlineFormat(TextElement element)
        {
            Bold = Local(element, TextElement.FontWeightProperty) is FontWeight w ? w > FontWeights.Normal : null;
            Italic = Local(element, TextElement.FontStyleProperty) is FontStyle s ? s != FontStyles.Normal : null;

            Underline = element switch
            {
                Inline inline when Local(inline, Inline.TextDecorationsProperty) is TextDecorationCollection d
                    => d.Count > 0,
                Paragraph p when Local(p, Paragraph.TextDecorationsProperty) is TextDecorationCollection pd
                    => pd.Count > 0,
                _ => null,
            };

            FontFamily = Local(element, TextElement.FontFamilyProperty) is FontFamily f ? f.Source : null;
            FontSizePx = Local(element, TextElement.FontSizeProperty) is double size ? size : null;
            Color = Local(element, TextElement.ForegroundProperty) is SolidColorBrush b ? ToHex(b.Color) : null;
        }

        private static object? Local(DependencyObject o, DependencyProperty p)
        {
            var v = o.ReadLocalValue(p);
            return v == DependencyProperty.UnsetValue ? null : v;
        }

        /// <summary>Values set here win; anything unset falls back to <paramref name="parent"/>.</summary>
        public InlineFormat MergedOver(InlineFormat parent) => new(
            Bold ?? parent.Bold,
            Italic ?? parent.Italic,
            Underline ?? parent.Underline,
            FontFamily ?? parent.FontFamily,
            FontSizePx ?? parent.FontSizePx,
            Color ?? parent.Color);

        /// <summary>Drops the colour and underline a Hyperlink gets from its control template.</summary>
        public InlineFormat WithoutLinkChrome() => new(Bold, Italic, null, FontFamily, FontSizePx, null);
    }

    private static string ToHex(Color c) => $"#{c.R:X2}{c.G:X2}{c.B:X2}";

    // ---------- HTML -> FlowDocument ----------

    private static readonly Regex RxToken = new(
        @"<\s*(?<closing>/)?\s*(?<name>[a-zA-Z][a-zA-Z0-9]*)(?<attrs>(?:""[^""]*""|'[^']*'|[^>])*?)\s*/?\s*>",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex RxAttr = new(
        @"(?<name>[a-zA-Z_:][-a-zA-Z0-9_:.]*)\s*=\s*(?:""(?<v1>[^""]*)""|'(?<v2>[^']*)'|(?<v3>[^\s""'>]+))",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    /// <summary>
    /// Replaces the document's content with the parsed HTML. The input is
    /// sanitized first, so this only ever sees allowlisted, well-nested markup.
    /// </summary>
    public static void Load(FlowDocument document, string? html)
    {
        document.Blocks.Clear();

        var fragment = ReplyHtml.Sanitize(html);
        if (fragment.Length == 0)
        {
            document.Blocks.Add(new Paragraph());
            return;
        }

        var blocks = new Stack<BlockCollection>();
        blocks.Push(document.Blocks);
        var lists = new Stack<List>();
        var spans = new Stack<Span>();
        Paragraph? para = null;

        Paragraph EnsureParagraph()
        {
            if (para is null)
            {
                para = new Paragraph { Margin = new Thickness(0) };
                blocks.Peek().Add(para);
            }
            return para;
        }

        void AddInline(Inline inline)
        {
            if (spans.Count > 0) spans.Peek().Inlines.Add(inline);
            else EnsureParagraph().Inlines.Add(inline);
        }

        void PushSpan(Span span)
        {
            AddInline(span);
            spans.Push(span);
        }

        var pos = 0;
        foreach (Match m in RxToken.Matches(fragment))
        {
            AddText(fragment.Substring(pos, m.Index - pos));
            pos = m.Index + m.Length;

            var name = m.Groups["name"].Value.ToLowerInvariant();
            var closing = m.Groups["closing"].Success;
            var attrs = m.Groups["attrs"].Value;

            switch (name)
            {
                case "br":
                    if (!closing) AddInline(new LineBreak());
                    break;

                case "p":
                case "div":
                    // Paragraph boundaries also terminate any inline formatting
                    // still open, matching how a browser reflows the content.
                    para = null;
                    spans.Clear();
                    if (!closing) EnsureParagraph();
                    break;

                case "ul":
                case "ol":
                    para = null;
                    spans.Clear();
                    if (closing)
                    {
                        if (lists.Count > 0) lists.Pop();
                    }
                    else
                    {
                        var list = new List
                        {
                            MarkerStyle = name == "ol" ? TextMarkerStyle.Decimal : TextMarkerStyle.Disc,
                            Margin = new Thickness(0),
                            Padding = new Thickness(20, 0, 0, 0),
                        };
                        blocks.Peek().Add(list);
                        lists.Push(list);
                    }
                    break;

                case "li":
                    para = null;
                    spans.Clear();
                    if (closing)
                    {
                        if (blocks.Count > 1) blocks.Pop();
                    }
                    else if (lists.Count > 0)
                    {
                        var item = new ListItem();
                        lists.Peek().ListItems.Add(item);
                        blocks.Push(item.Blocks);
                    }
                    break;

                case "a":
                    if (closing) { if (spans.Count > 0) spans.Pop(); }
                    else PushSpan(BuildHyperlink(attrs));
                    break;

                case "b":
                case "strong":
                    if (closing) { if (spans.Count > 0) spans.Pop(); }
                    else PushSpan(new Span { FontWeight = FontWeights.Bold });
                    break;

                case "i":
                case "em":
                    if (closing) { if (spans.Count > 0) spans.Pop(); }
                    else PushSpan(new Span { FontStyle = FontStyles.Italic });
                    break;

                case "u":
                    if (closing) { if (spans.Count > 0) spans.Pop(); }
                    else PushSpan(new Span { TextDecorations = TextDecorations.Underline });
                    break;

                case "span":
                case "font":
                    if (closing) { if (spans.Count > 0) spans.Pop(); }
                    else PushSpan(BuildStyledSpan(name, attrs));
                    break;
            }
        }

        AddText(fragment.Substring(pos));

        if (document.Blocks.Count == 0)
            document.Blocks.Add(new Paragraph());

        void AddText(string raw)
        {
            if (raw.Length == 0) return;
            var text = WebUtility.HtmlDecode(raw);
            // Markup indentation is not content: a run of pure whitespace
            // between tags would otherwise show up as a stray space.
            if (text.Trim().Length == 0 && spans.Count == 0 && para is null) return;
            AddInline(new Run(text.Replace("\r", string.Empty).Replace("\n", " ")));
        }
    }

    private static Hyperlink BuildHyperlink(string attrs)
    {
        var link = new Hyperlink();
        foreach (Match a in RxAttr.Matches(attrs))
        {
            if (!string.Equals(a.Groups["name"].Value, "href", StringComparison.OrdinalIgnoreCase)) continue;
            var raw = a.Groups["v1"].Success ? a.Groups["v1"].Value
                    : a.Groups["v2"].Success ? a.Groups["v2"].Value
                    : a.Groups["v3"].Value;
            var safe = ReplyHtml.SafeUrl(WebUtility.HtmlDecode(raw));
            if (safe != null && Uri.TryCreate(safe, UriKind.Absolute, out var uri))
            {
                link.NavigateUri = uri;
                link.ToolTip = safe;
            }
        }
        return link;
    }

    private static Span BuildStyledSpan(string tag, string attrs)
    {
        var span = new Span();
        foreach (Match a in RxAttr.Matches(attrs))
        {
            var name = a.Groups["name"].Value.ToLowerInvariant();
            var value = WebUtility.HtmlDecode(
                a.Groups["v1"].Success ? a.Groups["v1"].Value
                : a.Groups["v2"].Success ? a.Groups["v2"].Value
                : a.Groups["v3"].Value);

            if (name == "style") ApplyCss(span, value);
            else if (tag == "font" && name == "color") ApplyColor(span, value);
            else if (tag == "font" && name == "face") span.FontFamily = new FontFamily(value);
        }
        return span;
    }

    private static void ApplyCss(Span span, string css)
    {
        foreach (var decl in css.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var idx = decl.IndexOf(':');
            if (idx <= 0) continue;
            var prop = decl.Substring(0, idx).Trim().ToLowerInvariant();
            var val = decl.Substring(idx + 1).Trim();
            if (val.Length == 0) continue;

            switch (prop)
            {
                case "font-family":
                    // "Segoe UI, Arial" — WPF wants a single family here.
                    var first = val.Split(',')[0].Trim().Trim('\'', '"');
                    if (first.Length > 0) span.FontFamily = new FontFamily(first);
                    break;

                case "font-size":
                    if (TryParseCssSize(val, out var px)) span.FontSize = px;
                    break;

                case "color":
                    ApplyColor(span, val);
                    break;

                case "font-weight":
                    span.FontWeight = val.Equals("normal", StringComparison.OrdinalIgnoreCase) || val == "400"
                        ? FontWeights.Normal : FontWeights.Bold;
                    break;

                case "font-style":
                    span.FontStyle = val.Equals("normal", StringComparison.OrdinalIgnoreCase)
                        ? FontStyles.Normal : FontStyles.Italic;
                    break;

                case "text-decoration":
                    if (val.IndexOf("underline", StringComparison.OrdinalIgnoreCase) >= 0)
                        span.TextDecorations = TextDecorations.Underline;
                    else if (val.IndexOf("line-through", StringComparison.OrdinalIgnoreCase) >= 0)
                        span.TextDecorations = TextDecorations.Strikethrough;
                    break;
            }
        }
    }

    private static void ApplyColor(Span span, string value)
    {
        try
        {
            if (ColorConverter.ConvertFromString(value.Trim()) is Color c)
                span.Foreground = new SolidColorBrush(c);
        }
        catch (FormatException) { /* unparseable colour: keep the inherited one */ }
        catch (NotSupportedException) { /* ditto */ }
    }

    private static bool TryParseCssSize(string value, out double px)
    {
        px = 0;
        var m = Regex.Match(value.Trim(), @"^(?<n>[0-9]+(\.[0-9]+)?)\s*(?<u>pt|px|em|rem|%)?$", RegexOptions.CultureInvariant);
        if (!m.Success) return false;
        if (!double.TryParse(m.Groups["n"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var n)) return false;

        px = m.Groups["u"].Value.ToLowerInvariant() switch
        {
            "px" => n,
            "em" or "rem" => n * 14 * PxPerPt,
            "%" => n / 100 * 14 * PxPerPt,
            _ => n * PxPerPt, // pt, and the CSS default when a bare number slips through
        };
        return px is > 1 and < 400;
    }
}
