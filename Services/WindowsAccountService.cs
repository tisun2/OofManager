using System.Runtime.InteropServices;
using System.Text;

namespace OofManager.Wpf.Services;

public interface IWindowsAccountService
{
    /// <summary>
    /// Returns the UPN of the currently logged-in Windows user, or null if the
    /// account has no UPN (e.g. local account, workgroup machine). Useful as a
    /// hint to MSAL/WAM so an Entra-joined device can perform a silent SSO
    /// without ever prompting the user — the first launch on a corporate
    /// machine becomes a no-click sign-in.
    /// </summary>
    string? TryGetCurrentUserUpn();
}

public sealed class WindowsAccountService : IWindowsAccountService
{
    // NameUserPrincipal == 8 — see EXTENDED_NAME_FORMAT in secext.h.
    // GetUserNameEx returns the AAD/AD UPN (e.g. alice@contoso.com) for the
    // currently logged-in Windows user, even on Entra-joined or hybrid-joined
    // machines where the local NT account name is the SID-based AzureAD\alice
    // form. For workgroup / local-only accounts it fails (last error
    // ERROR_NONE_MAPPED), which we silently swallow.
    private const int NameUserPrincipal = 8;

    [DllImport("secur32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool GetUserNameEx(int nameFormat, StringBuilder lpNameBuffer, ref uint nSize);

    public string? TryGetCurrentUserUpn()
    {
        try
        {
            // 1024 chars is the documented max length for a UPN; larger than
            // any realistic value but still trivial to allocate once.
            var size = (uint)1024;
            var sb = new StringBuilder((int)size);
            if (GetUserNameEx(NameUserPrincipal, sb, ref size))
            {
                var upn = sb.ToString();
                // Sanity check: a real UPN always contains an '@'.
                return upn.Contains('@') ? upn : null;
            }
        }
        catch
        {
            // P/Invoke can throw on extremely locked-down Windows installs
            // (e.g. server core without secur32 in the image). Treat as "no
            // UPN available" — falls back to the existing Sign In button.
        }
        return null;
    }
}
