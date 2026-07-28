using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace OofManager.Wpf.Services;

/// <summary>
/// Single home for every transformation applied to an out-of-office reply body.
///
/// Exchange stores auto-reply messages as HTML, so the app's reply fields hold
/// HTML too. Everything that crosses a trust boundary funnels through here:
/// HTML arriving from Exchange, HTML leaving the rich-text editor, and the
/// flattened text embedded into the generated setup guide.
///
/// Two shapes are in play and the names keep them apart:
///   * <b>fragment</b>  — body-level markup only, what the editor round-trips.
///   * <b>document</b>  — the <c>&lt;html&gt;&lt;body&gt;…</c> envelope Exchange
///     and the Power Automate connector expect on the wire.
/// </summary>
public static class ReplyHtml
{
    // Compiled once per process; LoadAsync runs several of these on every
    // refresh (internal + external reply), so caching saves measurable CPU.
    private static readonly RegexOptions CompiledFast = RegexOptions.Compiled | RegexOptions.CultureInvariant;

    // ---------- HTML -> plain text (moved verbatim from ExchangeService) ----------

    private static readonly Regex RxComments = new(@"<!--.*?-->", CompiledFast | RegexOptions.Singleline);
    private static readonly Regex RxScriptStyle = new(@"<\s*(script|style)[^>]*>.*?<\s*/\s*\1\s*>", CompiledFast | RegexOptions.IgnoreCase | RegexOptions.Singleline);
    private static readonly Regex RxBr = new(@"<\s*br\s*/?\s*>", CompiledFast | RegexOptions.IgnoreCase);
    private static readonly Regex RxBlockClose = new(@"</\s*(p|div|li|tr|table|h[1-6])\s*>", CompiledFast | RegexOptions.IgnoreCase);
    private static readonly Regex RxCellClose = new(@"</\s*(td|th)\s*>", CompiledFast | RegexOptions.IgnoreCase);
    private static readonly Regex RxLiOpen = new(@"<\s*li[^>]*>", CompiledFast | RegexOptions.IgnoreCase);
    private static readonly Regex RxAnyTag = new(@"<[^>]+>", CompiledFast);
    private static readonly Regex RxBidi = new(@"[\u200E\u200F\u061C\u202A-\u202E\u2066-\u2069]", CompiledFast);
    private static readonly Regex RxZeroWidth = new(@"[\u200B-\u200D\uFEFF\u2060]", CompiledFast);
    private static readonly Regex RxSoftHyphen = new(@"[\u00AD\u2028\u2029]", CompiledFast);
    private static readonly Regex RxUnicodeSpaces = new(@"[\u00A0\u1680\u2000-\u200A\u202F\u205F\u3000]", CompiledFast);
    private static readonly Regex RxRunOfSpaces = new(@"[ \f\v]+", CompiledFast);
    private static readonly Regex RxSpacesAroundNl = new(@" *\n *", CompiledFast);
    private static readonly Regex RxBlankLines = new(@"\n{3,}", CompiledFast);

    /// <summary>
    /// Flattens HTML to readable plain text. Used for previews and for the
    /// generated setup guide, which must never carry live markup.
    /// </summary>
    public static string ToPlainText(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var text = value!;
        text = RxComments.Replace(text, string.Empty);
        text = RxScriptStyle.Replace(text, string.Empty);
        text = RxBr.Replace(text, "\n");
        text = RxBlockClose.Replace(text, "\n");
        text = RxCellClose.Replace(text, " ");
        text = RxLiOpen.Replace(text, "- ");
        text = RxAnyTag.Replace(text, string.Empty);
        text = WebUtility.HtmlDecode(text);

        text = RxBidi.Replace(text, string.Empty);
        text = RxZeroWidth.Replace(text, string.Empty);
        text = RxSoftHyphen.Replace(text, string.Empty);

        text = RxUnicodeSpaces.Replace(text, " ");
        text = text.Replace("\t", " ");
        text = text.Replace("\r\n", "\n").Replace("\r", "\n");
        text = RxRunOfSpaces.Replace(text, " ");
        text = RxSpacesAroundNl.Replace(text, "\n");
        text = RxBlankLines.Replace(text, "\n\n");
        return text.Trim();
    }

    // ---------- plain text -> HTML ----------

    /// <summary>
    /// Escapes plain text and turns newlines into &lt;br&gt;, producing a body
    /// fragment. This is the legacy-compatibility path for replies saved before
    /// the editor became rich-text.
    /// </summary>
    public static string FragmentFromPlainText(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var encoded = WebUtility.HtmlEncode(value!.Trim());
        encoded = encoded.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "<br>");
        return encoded;
    }

    // ---------- envelope handling ----------

    private static readonly Regex RxBodyOpen = new(@"<\s*body\b[^>]*>", CompiledFast | RegexOptions.IgnoreCase);
    private static readonly Regex RxBodyClose = new(@"<\s*/\s*body\s*>", CompiledFast | RegexOptions.IgnoreCase);
    private static readonly Regex RxLooksLikeHtml = new(@"<\s*/?\s*[a-zA-Z]", CompiledFast);

    /// <summary>
    /// True when the stored value already carries markup. Replies written by
    /// builds before the rich-text editor are bare plain text, and this is what
    /// tells the two apart — no schema change or pref migration required.
    /// </summary>
    public static bool LooksLikeHtml(string? value)
        => !string.IsNullOrWhiteSpace(value) && RxLooksLikeHtml.IsMatch(value!);

    /// <summary>
    /// Strips the <c>&lt;html&gt;/&lt;body&gt;</c> envelope, returning just the
    /// body-level markup. Exchange hands back a full document; the editor only
    /// wants what is inside it.
    /// </summary>
    private static string StripEnvelope(string html)
    {
        var open = RxBodyOpen.Match(html);
        if (open.Success)
        {
            var start = open.Index + open.Length;
            var close = RxBodyClose.Match(html, start);
            var end = close.Success ? close.Index : html.Length;
            return html.Substring(start, end - start);
        }
        return html;
    }

    /// <summary>
    /// Wraps a body fragment in the minimal HTML envelope Exchange and the
    /// Office 365 connector expect. Without it Outlook renders the reply as a
    /// single unformatted line.
    /// </summary>
    private static string ToDocument(string fragment)
        => string.IsNullOrWhiteSpace(fragment) ? string.Empty : $"<html><body>{fragment}</body></html>";

    // ---------- public conversions ----------

    /// <summary>
    /// Converts a stored reply into a sanitized body fragment for the editor.
    /// Accepts both rich HTML and legacy plain text.
    /// </summary>
    public static string ToEditorFragment(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        return LooksLikeHtml(value)
            ? Sanitize(StripEnvelope(value!))
            : FragmentFromPlainText(value);
    }

    /// <summary>
    /// Converts a stored reply into the sanitized full HTML document that goes
    /// on the wire to Exchange or into a Power Automate flow definition.
    /// Accepts both rich HTML and legacy plain text, so every call site can use
    /// this without knowing which era the value came from.
    /// </summary>
    public static string EnsureHtml(string? value)
        => ToDocument(ToEditorFragment(value));

    // ---------- comparison ----------

    private static readonly Regex RxBetweenTags = new(@">\s+<", CompiledFast);
    private static readonly Regex RxWhitespaceRun = new(@"\s+", CompiledFast);

    /// <summary>
    /// Canonical form for equality checks between a locally-authored reply and
    /// one read back out of a deployed cloud flow. Collapses insignificant
    /// whitespace so cosmetic differences don't read as real drift.
    /// </summary>
    public static string Normalize(string? html)
    {
        if (string.IsNullOrWhiteSpace(html)) return string.Empty;
        var s = RxBetweenTags.Replace(html!, "><");
        s = RxWhitespaceRun.Replace(s, " ");
        return s.Trim();
    }

    // ---------- sanitizer ----------

    // Element allowlist. Anything outside this set is unwrapped (its text
    // survives, the tag does not) rather than dropped wholesale, so stripping
    // an unexpected <table> doesn't silently delete the user's words.
    private static readonly HashSet<string> AllowedTags = new(StringComparer.OrdinalIgnoreCase)
    {
        "b", "strong", "i", "em", "u", "span", "font", "a", "br", "p", "div", "ul", "ol", "li",
    };

    // Elements whose *content* is discarded along with the tag. Unwrapping
    // these would splice script source or CSS into the visible message.
    private static readonly Regex RxDropWithContent = new(
        @"<\s*(script|style|head|title|iframe|object|embed|noscript|template|svg|math)\b[^>]*>.*?<\s*/\s*\1\s*>",
        CompiledFast | RegexOptions.IgnoreCase | RegexOptions.Singleline);

    // Unpaired/void elements that must never survive even without a closing tag.
    private static readonly Regex RxDropStandalone = new(
        @"<\s*/?\s*(script|style|head|title|iframe|object|embed|noscript|template|svg|math|img|input|form|base|meta|link|button|textarea|select|option)\b[^>]*>",
        CompiledFast | RegexOptions.IgnoreCase);

    private static readonly Regex RxVoidTags = new(@"^(br)$", CompiledFast | RegexOptions.IgnoreCase);

    private static readonly Regex RxTag = new(
        @"<\s*(?<closing>/)?\s*(?<name>[a-zA-Z][a-zA-Z0-9]*)(?<attrs>(?:""[^""]*""|'[^']*'|[^>])*?)\s*(?<self>/)?\s*>",
        CompiledFast);

    private static readonly Regex RxAttr = new(
        @"(?<name>[a-zA-Z_:][-a-zA-Z0-9_:.]*)\s*=\s*(?:""(?<v1>[^""]*)""|'(?<v2>[^']*)'|(?<v3>[^\s""'>]+))",
        CompiledFast);

    /// <summary>
    /// Reduces arbitrary HTML to the small, safe subset the editor understands.
    /// Runs on both ingress (Exchange -> editor) and egress (editor -> Exchange,
    /// flow package), so untrusted markup can neither reach the app nor be
    /// relayed onward to a recipient's mail client.
    /// </summary>
    public static string Sanitize(string? html)
    {
        if (string.IsNullOrWhiteSpace(html)) return string.Empty;

        var input = RxComments.Replace(html!, string.Empty);
        input = RxDropWithContent.Replace(input, string.Empty);
        input = RxDropStandalone.Replace(input, string.Empty);

        var sb = new StringBuilder(input.Length);
        var open = new Stack<string>();
        var pos = 0;

        foreach (Match m in RxTag.Matches(input))
        {
            AppendText(sb, input.Substring(pos, m.Index - pos));
            pos = m.Index + m.Length;

            var name = m.Groups["name"].Value.ToLowerInvariant();
            if (!AllowedTags.Contains(name))
                continue; // unwrap: the tag vanishes, its text keeps flowing

            var isClosing = m.Groups["closing"].Success;
            var isVoid = RxVoidTags.IsMatch(name);

            if (isClosing)
            {
                // Ignore a stray close with no matching open, and unwind any
                // elements the source left dangling in between.
                if (isVoid || !open.Contains(name)) continue;
                string popped;
                do
                {
                    popped = open.Pop();
                    sb.Append("</").Append(popped).Append('>');
                } while (!string.Equals(popped, name, StringComparison.Ordinal));
                continue;
            }

            if (isVoid)
            {
                sb.Append("<").Append(name).Append(">");
                continue;
            }

            var attrs = SanitizeAttributes(name, m.Groups["attrs"].Value);
            // A link whose href didn't survive validation is no longer a link;
            // unwrap it so the anchor text still reads normally.
            if (name == "a" && attrs.Length == 0)
                continue;

            if (m.Groups["self"].Success)
            {
                sb.Append('<').Append(name).Append(attrs).Append("></").Append(name).Append('>');
                continue;
            }

            sb.Append('<').Append(name).Append(attrs).Append('>');
            open.Push(name);
        }

        AppendText(sb, input.Substring(pos));

        while (open.Count > 0)
            sb.Append("</").Append(open.Pop()).Append('>');

        return sb.ToString();
    }

    // Text nodes are decoded then re-encoded so entity-smuggled markup
    // ("&lt;script&gt;") stays inert and stray '<' characters can't reopen the
    // parser downstream.
    private static void AppendText(StringBuilder sb, string raw)
    {
        if (raw.Length == 0) return;
        sb.Append(WebUtility.HtmlEncode(WebUtility.HtmlDecode(raw)));
    }

    private static string SanitizeAttributes(string tag, string rawAttrs)
    {
        if (string.IsNullOrWhiteSpace(rawAttrs)) return string.Empty;

        var sb = new StringBuilder();
        foreach (Match a in RxAttr.Matches(rawAttrs))
        {
            var name = a.Groups["name"].Value.ToLowerInvariant();
            var value = a.Groups["v1"].Success ? a.Groups["v1"].Value
                      : a.Groups["v2"].Success ? a.Groups["v2"].Value
                      : a.Groups["v3"].Value;

            // Every on* handler is dropped by omission — nothing below accepts one.
            switch (name)
            {
                case "href" when tag == "a":
                    var url = SafeUrl(value);
                    if (url != null)
                        sb.Append(" href=\"").Append(WebUtility.HtmlEncode(url)).Append('"');
                    break;

                case "style" when tag is "span" or "p" or "div" or "li" or "ul" or "ol" or "a" or "b" or "i" or "u" or "strong" or "em" or "font":
                    var css = SafeCss(value);
                    if (css.Length > 0)
                        sb.Append(" style=\"").Append(WebUtility.HtmlEncode(css)).Append('"');
                    break;

                case "color" when tag == "font":
                    if (RxCssColor.IsMatch(value.Trim()))
                        sb.Append(" color=\"").Append(WebUtility.HtmlEncode(value.Trim())).Append('"');
                    break;

                case "face" when tag == "font":
                    if (RxCssFontFamily.IsMatch(value.Trim()))
                        sb.Append(" face=\"").Append(WebUtility.HtmlEncode(value.Trim())).Append('"');
                    break;

                case "size" when tag == "font":
                    if (RxFontTagSize.IsMatch(value.Trim()))
                        sb.Append(" size=\"").Append(WebUtility.HtmlEncode(value.Trim())).Append('"');
                    break;
            }
        }
        return sb.ToString();
    }

    // ---------- URL validation ----------

    // Browsers and mail clients ignore control characters and whitespace inside
    // a URL, so "java&#9;script:alert(1)" runs. Strip them before the scheme is
    // read, otherwise the check below inspects a string nobody will execute.
    private static readonly Regex RxUrlNoise = new(@"[\x00-\x20\x7F\u00A0\u200B-\u200D\uFEFF\u2060]", CompiledFast);
    private static readonly Regex RxScheme = new(@"^(?<s>[a-zA-Z][a-zA-Z0-9+.\-]*):", CompiledFast);

    private static readonly HashSet<string> AllowedSchemes = new(StringComparer.OrdinalIgnoreCase)
    {
        "http", "https", "mailto",
    };

    /// <summary>
    /// Returns the URL when it is safe to emit, otherwise null. Only absolute
    /// http/https/mailto links are accepted: relative URLs are meaningless once
    /// the reply is rendered in someone else's mail client, and every other
    /// scheme (javascript:, data:, vbscript:, file:) is an execution vector.
    /// </summary>
    public static string? SafeUrl(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;

        var decoded = WebUtility.HtmlDecode(value!);
        var probe = RxUrlNoise.Replace(decoded, string.Empty);
        if (probe.Length == 0) return null;

        var m = RxScheme.Match(probe);
        if (!m.Success || !AllowedSchemes.Contains(m.Groups["s"].Value))
            return null;

        // Emit the noise-free form so the value that ships is the same one that
        // was validated.
        return probe;
    }

    /// <summary>
    /// Best-effort normalization of user input from the link dialog: bare hosts
    /// like "example.com" become "https://example.com", bare addresses become
    /// "mailto:". Returns null when the result still isn't a safe URL.
    /// </summary>
    public static string? NormalizeUserUrl(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var trimmed = RxUrlNoise.Replace(value!, string.Empty);
        if (trimmed.Length == 0) return null;

        if (!RxScheme.IsMatch(trimmed))
        {
            trimmed = trimmed.IndexOf('@') > 0 && trimmed.IndexOf('/') < 0
                ? "mailto:" + trimmed
                : "https://" + trimmed;
        }
        return SafeUrl(trimmed);
    }

    // ---------- CSS validation ----------

    private static readonly Regex RxCssFontFamily = new(@"^[a-zA-Z0-9 ,'""\-]{1,120}$", CompiledFast);
    private static readonly Regex RxCssFontSize = new(@"^[0-9]{1,3}(\.[0-9]{1,2})?(pt|px|em|rem|%)$", CompiledFast);
    private static readonly Regex RxCssColor = new(@"^(#[0-9a-fA-F]{3,8}|[a-zA-Z]{1,24}|rgba?\([0-9.,%\s]{1,40}\))$", CompiledFast);
    private static readonly Regex RxCssFontWeight = new(@"^(normal|bold|bolder|lighter|[1-9]00)$", CompiledFast | RegexOptions.IgnoreCase);
    private static readonly Regex RxCssFontStyle = new(@"^(normal|italic|oblique)$", CompiledFast | RegexOptions.IgnoreCase);
    private static readonly Regex RxCssDecoration = new(@"^(none|underline|line-through|overline)( (none|underline|line-through|overline))*$", CompiledFast | RegexOptions.IgnoreCase);
    private static readonly Regex RxFontTagSize = new(@"^[+\-]?[1-7]$", CompiledFast);

    /// <summary>
    /// Rebuilds a style attribute from only the declarations the editor can
    /// round-trip, each matched against a value pattern. Building the result
    /// from scratch — rather than blocking known-bad substrings — means
    /// url(), expression() and every other payload fail by simply not matching.
    /// </summary>
    private static string SafeCss(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;

        var sb = new StringBuilder();
        foreach (var decl in WebUtility.HtmlDecode(value!).Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var idx = decl.IndexOf(':');
            if (idx <= 0) continue;

            var prop = decl.Substring(0, idx).Trim().ToLowerInvariant();
            var val = decl.Substring(idx + 1).Trim().TrimEnd('!').Trim();
            if (val.Length == 0) continue;

            var ok = prop switch
            {
                "font-family" => RxCssFontFamily.IsMatch(val),
                "font-size" => RxCssFontSize.IsMatch(val),
                "color" => RxCssColor.IsMatch(val),
                "font-weight" => RxCssFontWeight.IsMatch(val),
                "font-style" => RxCssFontStyle.IsMatch(val),
                "text-decoration" => RxCssDecoration.IsMatch(val),
                _ => false,
            };
            if (!ok) continue;

            if (sb.Length > 0) sb.Append(' ');
            sb.Append(prop).Append(':').Append(val).Append(';');
        }
        return sb.ToString();
    }
}
