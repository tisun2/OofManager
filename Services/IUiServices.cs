namespace OofManager.Wpf.Services;

public interface IDialogService
{
    Task<bool> ConfirmAsync(string title, string message, string accept = "OK", string cancel = "Cancel");
    Task AlertAsync(string title, string message, string close = "OK");
    Task<string?> PromptAsync(string title, string message, string accept = "OK", string cancel = "Cancel", string? placeholder = null);
}

public interface INavigationService
{
    void NavigateToMain();
    /// <summary>
    /// Navigates back to the LoginPage. When <paramref name="forceAccountPicker"/>
    /// is true, also kicks off an interactive WAM sign-in with no UPN hint so
    /// the user can choose a different Microsoft 365 account (used by the
    /// Switch Account button on the main page).
    /// </summary>
    void NavigateToLogin(bool forceAccountPicker = false);
}

public interface IPreferencesService
{
    bool GetBool(string key, bool defaultValue);
    int GetInt(string key, int defaultValue);
    string? GetString(string key, string? defaultValue = null);
    void Set(string key, bool value);
    void Set(string key, int value);
    void Set(string key, string? value);
    /// <summary>
    /// Suppresses immediate disk writes from Set() calls. Call Resume() to flush
    /// once at the end. Use to batch many sets into a single write.
    /// </summary>
    IDisposable BeginBatch();
}
