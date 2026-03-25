using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Animation;
using Drawing = System.Drawing;
using Forms = System.Windows.Forms;

namespace AMTool;

public partial class MainWindow : Window
{
    private const int WM_HOTKEY = 0x0312;
    private const int WH_MOUSE_LL = 14;
    private const int WM_LBUTTONDOWN = 0x0201;
    private const int WM_RBUTTONDOWN = 0x0204;
    private const int WM_MBUTTONDOWN = 0x0207;
    private const int WM_XBUTTONDOWN = 0x020B;
    private const int HotkeyIdToggleWindow = 9001;
    private const int VK_SHIFT = 0x10;
    private const int VK_CONTROL = 0x11;
    private const int VK_MENU = 0x12;
    private const int VK_LWIN = 0x5B;
    private const int VK_RWIN = 0x5C;
    private const ushort XButton1 = 0x0001;
    private const ushort XButton2 = 0x0002;
    private const uint MOD_CONTROL = 0x0002;
    private const uint MOD_SHIFT = 0x0004;
    private const uint MOD_ALT = 0x0001;
    private const uint MOD_WIN = 0x0008;
    private const uint MOD_NOREPEAT = 0x4000;
    private const double BottomMargin = 24;
    private const int MaxVisibleOrbitTiles = 5;
    private const double OrbitTileSize = 88;
    private const double OrbitTileCornerRadius = OrbitTileSize / 2;
    private const double ShortcutIconSize = 44;
    private const double ShortcutIconShellSize = 64;
    private const double OrbitEntryScale = 0.16;
    private const double OrbitExitScale = 0.18;
    private const double OrbitEntryAngle = 28;
    private const double OrbitExitAngle = 24;
    private const int OrbitEntryDurationMs = 320;
    private const int OrbitExitDurationMs = 220;
    private const int OrbitScrollDurationMs = 200;
    private const int OrbitScrollFastDurationMs = OrbitScrollDurationMs / 3;
    private const int OrbitScrollBurstThresholdMs = 180;
    private const int OrbitSwapDurationMs = 240;
    private const int OrbitAnimationStaggerMs = 34;
    private const int OrbitHoverDurationMs = 140;
    private const int HiddenIdleCleanupDelayMs = 12000;
    private const double DraggedTileScale = 1.08;
    private const double DragTargetScale = 1.04;
    private const double HoverTileScale = 1.06;
    private const string AddOrbitItemKey = "__add__";
    private const string AutoStartRegistryValueName = "AMTool";

    private static readonly System.Windows.Point[] OrbitSlotPositions =
    [
        new(28, 152),
        new(172, 52),
        new(346, 12),
        new(520, 52),
        new(664, 152)
    ];
    private static readonly System.Windows.Media.Brush ShortcutTileBackgroundBrush = CreateFrozenBrush(42, 244, 247, 251);
    private static readonly System.Windows.Media.Brush ShortcutTileBorderBrush = CreateFrozenBrush(92, 244, 247, 251);
    private static readonly System.Windows.Media.Brush AddTileBackgroundBrush = CreateFrozenBrush(28, 244, 247, 251);
    private static readonly System.Windows.Media.Brush AddTileBorderBrush = CreateFrozenBrush(118, 244, 247, 251);
    private static readonly BitmapCache OrbitTileBitmapCache = CreateFrozenBitmapCache();
    private static readonly IEasingFunction OrbitEaseOut = CreateFrozenCubicEase(EasingMode.EaseOut);
    private static readonly IEasingFunction OrbitEaseIn = CreateFrozenCubicEase(EasingMode.EaseIn);
    private static readonly BitmapSizeOptions ShortcutIconBitmapSizeOptions =
        BitmapSizeOptions.FromWidthAndHeight((int)ShortcutIconSize, (int)ShortcutIconSize);

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private readonly List<AppShortcutEntry> _shortcuts = [];
    private readonly Dictionary<string, ImageSource?> _iconCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly string _appDataDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "AMTool");
    private readonly string _storagePath;
    private readonly string _settingsPath;

    private HwndSource? _hwndSource;
    private LowLevelMouseProc? _mouseHookProc;
    private IntPtr _mouseHookHandle;
    private Forms.NotifyIcon? _trayIcon;
    private Drawing.Icon? _applicationTrayIcon;
    private bool _isExitRequested;
    private bool _isHotkeyRegistered;
    private bool _isVisibilityAnimationRunning;
    private bool _isScrollAnimationRunning;
    private bool _isShortcutDragActive;
    private bool _isHiddenIdleModeActive;
    private int _shortcutScrollIndex;
    private int _orbitScrollDirection = 1;
    private int _orbitScrollInputVersion;
    private DateTime _lastOrbitScrollInputUtc = DateTime.MinValue;
    private bool _isOrbitScrollBurstActive;
    private bool _isAutoStartEnabled;
    private bool _hasCompletedTutorial;
    private CancellationTokenSource? _hiddenIdleCleanupCancellation;
    private ModifierKeys _hotkeyModifiers = ModifierKeys.Control | ModifierKeys.Shift;
    private HotkeyInputKind _hotkeyInputKind = HotkeyInputKind.Keyboard;
    private Key _hotkeyKey = Key.Q;
    private HotkeyMouseButton _hotkeyMouseButton;
    private Border? _pressedShortcutTile;
    private Border? _dragTargetTile;
    private AppShortcutEntry? _pressedShortcut;
    private System.Windows.Point _dragStartPointerPosition;
    private System.Windows.Point _dragSourceSlotPosition;

    public MainWindow()
    {
        InitializeComponent();
        _storagePath = Path.Combine(_appDataDirectory, "shortcuts.json");
        _settingsPath = Path.Combine(_appDataDirectory, "settings.json");
        InitializeInfoBadgeInteractions();

        SourceInitialized += MainWindow_SourceInitialized;
        Loaded += MainWindow_Loaded;
        Deactivated += MainWindow_Deactivated;
        Closing += MainWindow_Closing;
        Closed += MainWindow_Closed;

        LoadSettings();
        LoadShortcuts();
    }

    private void InitializeInfoBadgeInteractions()
    {
        InfoBadge.Cursor = System.Windows.Input.Cursors.Hand;
        InfoBadge.MouseRightButtonUp += InfoBadge_MouseRightButtonUp;
        UpdateHotkeyUi();
    }

    private void LoadSettings()
    {
        if (!File.Exists(_settingsPath))
        {
            UpdateHotkeyUi();
            return;
        }

        try
        {
            string json = File.ReadAllText(_settingsPath);
            AppSettings? settings = JsonSerializer.Deserialize<AppSettings>(json);

            if (settings is not null
                && HotkeyUtilities.IsValidHotkey(
                    settings.HotkeyModifiers,
                    settings.HotkeyKey,
                    settings.HotkeyInputKind,
                    settings.HotkeyMouseButton))
            {
                _hotkeyModifiers = HotkeyUtilities.SanitizeModifiers(settings.HotkeyModifiers);
                _hotkeyInputKind = settings.HotkeyInputKind;
                _hotkeyKey = settings.HotkeyInputKind == HotkeyInputKind.Keyboard ? settings.HotkeyKey : Key.None;
                _hotkeyMouseButton = settings.HotkeyInputKind == HotkeyInputKind.MouseButton
                    ? settings.HotkeyMouseButton
                    : HotkeyMouseButton.None;
            }

            _isAutoStartEnabled = settings?.AutoStartEnabled == true;
            _hasCompletedTutorial = settings?.HasCompletedTutorial == true;
        }
        catch
        {
        }

        ApplyAutoStartSetting(_isAutoStartEnabled);
        UpdateHotkeyUi();
    }

    private void SaveSettings()
    {
        try
        {
            Directory.CreateDirectory(_appDataDirectory);

            string json = JsonSerializer.Serialize(
                new AppSettings
                {
                    HotkeyModifiers = _hotkeyModifiers,
                    HotkeyInputKind = _hotkeyInputKind,
                    HotkeyKey = _hotkeyKey,
                    HotkeyMouseButton = _hotkeyMouseButton,
                    AutoStartEnabled = _isAutoStartEnabled,
                    HasCompletedTutorial = _hasCompletedTutorial
                },
                JsonOptions);

            File.WriteAllText(_settingsPath, json);
        }
        catch
        {
            System.Windows.MessageBox.Show(
                "Nie udalo sie zapisac ustawien.",
                "AMTool",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private bool TryShowDialog(Window dialog, string failureMessage, out bool? result)
    {
        result = null;

        try
        {
            result = dialog.ShowDialog();
            return true;
        }
        catch
        {
            System.Windows.MessageBox.Show(
                failureMessage,
                "AMTool",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }
    }

    private bool TryShowFileDialog(Microsoft.Win32.OpenFileDialog dialog, out bool? result)
    {
        result = null;

        try
        {
            result = dialog.ShowDialog(this);
            return true;
        }
        catch
        {
            System.Windows.MessageBox.Show(
                "Nie udalo sie otworzyc okna wyboru pliku.",
                "AMTool",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }
    }

    private void UpdateHotkeyUi()
    {
        if (InfoBadge is not null)
        {
            UpdateInfoTooltip(_shortcuts.Count + 1);
        }
    }

    private string GetCurrentHotkeyText()
    {
        return HotkeyUtilities.FormatHotkey(
            _hotkeyModifiers,
            _hotkeyKey,
            _hotkeyInputKind,
            _hotkeyMouseButton);
    }

    private void MainWindow_SourceInitialized(object? sender, EventArgs e)
    {
        IntPtr hwnd = new WindowInteropHelper(this).Handle;

        _hwndSource = HwndSource.FromHwnd(hwnd);
        _hwndSource?.AddHook(WndProc);

        RegisterHotkey(hwnd);
        InitializeTrayIcon();
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        HideToTrayImmediate();

        if (!_hasCompletedTutorial)
        {
            await ShowFirstRunTutorialAsync();
        }
    }

    private async void MainWindow_Deactivated(object? sender, EventArgs e)
    {
        if (_isExitRequested || !IsVisible || _isVisibilityAnimationRunning)
        {
            return;
        }

        await Task.Yield();

        if (_isExitRequested || !IsVisible || _isVisibilityAnimationRunning)
        {
            return;
        }

        if (IsForegroundWindowOwnedByCurrentProcess())
        {
            return;
        }

        await HideToTrayAsync();
    }

    private void MainWindow_Closing(object? sender, CancelEventArgs e)
    {
        if (_isExitRequested)
        {
            return;
        }

        e.Cancel = true;
        _ = HideToTrayAsync();
    }

    private void MainWindow_Closed(object? sender, EventArgs e)
    {
        CancelHiddenIdleCleanup();

        IntPtr hwnd = new WindowInteropHelper(this).Handle;

        if (_isHotkeyRegistered)
        {
            UnregisterHotKey(hwnd, HotkeyIdToggleWindow);
        }

        if (_hwndSource is not null)
        {
            _hwndSource.RemoveHook(WndProc);
            _hwndSource = null;
        }

        if (_mouseHookHandle != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_mouseHookHandle);
            _mouseHookHandle = IntPtr.Zero;
            _mouseHookProc = null;
        }

        if (_trayIcon is not null)
        {
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            _trayIcon = null;
        }

        if (_applicationTrayIcon is not null)
        {
            _applicationTrayIcon.Dispose();
            _applicationTrayIcon = null;
        }
    }

    private void RegisterHotkey(IntPtr hwnd)
    {
        if (_hotkeyInputKind == HotkeyInputKind.MouseButton)
        {
            _isHotkeyRegistered = false;
            UpdateMouseHookState();
            return;
        }

        _isHotkeyRegistered = TryRegisterKeyboardHotkey(hwnd, _hotkeyModifiers, _hotkeyKey);

        if (!_isHotkeyRegistered)
        {
            System.Windows.MessageBox.Show(
                $"Nie udalo sie zarejestrowac skrotu {GetCurrentHotkeyText()}.",
                "AMTool",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private static uint ToNativeHotkeyModifiers(ModifierKeys modifiers)
    {
        ModifierKeys sanitizedModifiers = HotkeyUtilities.SanitizeModifiers(modifiers);
        uint nativeModifiers = 0;

        if (sanitizedModifiers.HasFlag(ModifierKeys.Control))
        {
            nativeModifiers |= MOD_CONTROL;
        }

        if (sanitizedModifiers.HasFlag(ModifierKeys.Shift))
        {
            nativeModifiers |= MOD_SHIFT;
        }

        if (sanitizedModifiers.HasFlag(ModifierKeys.Alt))
        {
            nativeModifiers |= MOD_ALT;
        }

        if (sanitizedModifiers.HasFlag(ModifierKeys.Windows))
        {
            nativeModifiers |= MOD_WIN;
        }

        return nativeModifiers | MOD_NOREPEAT;
    }

    private static bool TryRegisterKeyboardHotkey(IntPtr hwnd, ModifierKeys modifiers, Key key)
    {
        return RegisterHotKey(
            hwnd,
            HotkeyIdToggleWindow,
            ToNativeHotkeyModifiers(modifiers),
            (uint)KeyInterop.VirtualKeyFromKey(key));
    }

    private bool EnsureMouseHook()
    {
        if (_mouseHookHandle != IntPtr.Zero)
        {
            return true;
        }

        _mouseHookProc = MouseHookCallback;

        using Process currentProcess = Process.GetCurrentProcess();
        string? moduleName = currentProcess.MainModule?.ModuleName;
        IntPtr moduleHandle = string.IsNullOrWhiteSpace(moduleName)
            ? IntPtr.Zero
            : GetModuleHandle(moduleName);

        _mouseHookHandle = SetWindowsHookEx(WH_MOUSE_LL, _mouseHookProc, moduleHandle, 0);
        return _mouseHookHandle != IntPtr.Zero;
    }

    private void ReleaseMouseHook()
    {
        if (_mouseHookHandle == IntPtr.Zero)
        {
            return;
        }

        UnhookWindowsHookEx(_mouseHookHandle);
        _mouseHookHandle = IntPtr.Zero;
        _mouseHookProc = null;
    }

    private void UpdateMouseHookState()
    {
        bool shouldHookBeActive =
            !_isExitRequested
            && _hotkeyInputKind == HotkeyInputKind.MouseButton
            && !IsVisible;

        if (shouldHookBeActive)
        {
            _ = EnsureMouseHook();
            return;
        }

        ReleaseMouseHook();
    }

    private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0)
        {
            TryHandleMouseHotkey(wParam, lParam);
        }

        return CallNextHookEx(_mouseHookHandle, nCode, wParam, lParam);
    }

    private void TryHandleMouseHotkey(IntPtr wParam, IntPtr lParam)
    {
        if (_isExitRequested
            || _hotkeyInputKind != HotkeyInputKind.MouseButton
            || _hotkeyMouseButton == HotkeyMouseButton.None
            || IsForegroundWindowOwnedByCurrentProcess()
            || !TryGetMouseButtonFromHookMessage(wParam, lParam, out HotkeyMouseButton mouseButton)
            || mouseButton != _hotkeyMouseButton
            || HotkeyUtilities.SanitizeModifiers(_hotkeyModifiers) != GetCurrentPressedModifiers())
        {
            return;
        }

        _ = Dispatcher.BeginInvoke(new Action(ToggleWindowVisibility));
    }

    private static bool TryGetMouseButtonFromHookMessage(IntPtr wParam, IntPtr lParam, out HotkeyMouseButton mouseButton)
    {
        switch (wParam.ToInt32())
        {
            case WM_LBUTTONDOWN:
                mouseButton = HotkeyMouseButton.Left;
                return true;
            case WM_RBUTTONDOWN:
                mouseButton = HotkeyMouseButton.Right;
                return true;
            case WM_MBUTTONDOWN:
                mouseButton = HotkeyMouseButton.Middle;
                return true;
            case WM_XBUTTONDOWN:
                MSLLHOOKSTRUCT hookData = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);
                ushort xButton = (ushort)((hookData.mouseData >> 16) & 0xFFFF);
                mouseButton = xButton switch
                {
                    XButton1 => HotkeyMouseButton.XButton1,
                    XButton2 => HotkeyMouseButton.XButton2,
                    _ => HotkeyMouseButton.None
                };
                return mouseButton != HotkeyMouseButton.None;
            default:
                mouseButton = HotkeyMouseButton.None;
                return false;
        }
    }

    private static ModifierKeys GetCurrentPressedModifiers()
    {
        ModifierKeys modifiers = ModifierKeys.None;

        if (IsVirtualKeyPressed(VK_CONTROL))
        {
            modifiers |= ModifierKeys.Control;
        }

        if (IsVirtualKeyPressed(VK_SHIFT))
        {
            modifiers |= ModifierKeys.Shift;
        }

        if (IsVirtualKeyPressed(VK_MENU))
        {
            modifiers |= ModifierKeys.Alt;
        }

        if (IsVirtualKeyPressed(VK_LWIN) || IsVirtualKeyPressed(VK_RWIN))
        {
            modifiers |= ModifierKeys.Windows;
        }

        return HotkeyUtilities.SanitizeModifiers(modifiers);
    }

    private static bool IsVirtualKeyPressed(int virtualKey)
    {
        return (GetAsyncKeyState(virtualKey) & 0x8000) != 0;
    }

    private bool TryApplyHotkey(
        ModifierKeys newModifiers,
        Key newKey,
        HotkeyInputKind newInputKind,
        HotkeyMouseButton newMouseButton)
    {
        if (_hwndSource is null
            || !HotkeyUtilities.IsValidHotkey(newModifiers, newKey, newInputKind, newMouseButton))
        {
            return false;
        }

        IntPtr hwnd = new WindowInteropHelper(this).Handle;
        ModifierKeys previousModifiers = _hotkeyModifiers;
        HotkeyInputKind previousInputKind = _hotkeyInputKind;
        Key previousKey = _hotkeyKey;
        HotkeyMouseButton previousMouseButton = _hotkeyMouseButton;

        if (_isHotkeyRegistered)
        {
            UnregisterHotKey(hwnd, HotkeyIdToggleWindow);
            _isHotkeyRegistered = false;
        }

        _hotkeyModifiers = HotkeyUtilities.SanitizeModifiers(newModifiers);
        _hotkeyInputKind = newInputKind;
        _hotkeyKey = newInputKind == HotkeyInputKind.Keyboard ? newKey : Key.None;
        _hotkeyMouseButton = newInputKind == HotkeyInputKind.MouseButton
            ? newMouseButton
            : HotkeyMouseButton.None;

        bool hotkeyActivated = newInputKind switch
        {
            HotkeyInputKind.Keyboard => TryRegisterKeyboardHotkey(hwnd, _hotkeyModifiers, _hotkeyKey),
            HotkeyInputKind.MouseButton => EnsureMouseHook(),
            _ => false
        };
        _isHotkeyRegistered = newInputKind == HotkeyInputKind.Keyboard && hotkeyActivated;

        if (hotkeyActivated)
        {
            UpdateMouseHookState();
            SaveSettings();
            UpdateHotkeyUi();
            return true;
        }

        _hotkeyModifiers = previousModifiers;
        _hotkeyInputKind = previousInputKind;
        _hotkeyKey = previousKey;
        _hotkeyMouseButton = previousMouseButton;
        _isHotkeyRegistered = previousInputKind == HotkeyInputKind.Keyboard
            && TryRegisterKeyboardHotkey(hwnd, _hotkeyModifiers, _hotkeyKey);
        UpdateMouseHookState();

        string failureReason = newInputKind == HotkeyInputKind.MouseButton
            ? "Globalny nasluch myszy nie jest dostepny."
            : "Ten skrot moze byc juz zajety.";

        System.Windows.MessageBox.Show(
            $"Nie udalo sie ustawic skrotu {HotkeyUtilities.FormatHotkey(newModifiers, newKey, newInputKind, newMouseButton)}. {failureReason}",
            "AMTool",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        UpdateHotkeyUi();
        return false;
    }

    private void ApplyAutoStartSetting(bool enabled)
    {
        const string runKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";

        try
        {
            using Microsoft.Win32.RegistryKey? runKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(runKeyPath, writable: true);

            if (runKey is null)
            {
                return;
            }

            if (enabled)
            {
                string executablePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName ?? string.Empty;

                if (!string.IsNullOrWhiteSpace(executablePath))
                {
                    runKey.SetValue(AutoStartRegistryValueName, $"\"{executablePath}\"");
                }
            }
            else
            {
                runKey.DeleteValue(AutoStartRegistryValueName, throwOnMissingValue: false);
            }
        }
        catch
        {
            System.Windows.MessageBox.Show(
                "Nie udalo sie zaktualizowac ustawienia autostartu.",
                "AMTool",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void ApplyAndPersistAutoStartSetting(bool enabled)
    {
        _isAutoStartEnabled = enabled;
        ApplyAutoStartSetting(_isAutoStartEnabled);
        SaveSettings();
        UpdateHotkeyUi();
    }

    private void OpenSettingsDialog()
    {
        while (true)
        {
            SettingsDialog dialog;

            try
            {
                dialog = new SettingsDialog(
                    _hotkeyModifiers,
                    _hotkeyKey,
                    _hotkeyInputKind,
                    _hotkeyMouseButton,
                    _isAutoStartEnabled)
                {
                    Owner = this
                };
            }
            catch
            {
                System.Windows.MessageBox.Show(
                    "Nie udalo sie przygotowac okna ustawien.",
                    "AMTool",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (!TryShowDialog(dialog, "Nie udalo sie otworzyc okna ustawien.", out bool? dialogResult))
            {
                return;
            }

            if (dialogResult != true)
            {
                return;
            }

            bool hotkeyChanged =
                dialog.SelectedHotkeyModifiers != _hotkeyModifiers
                || dialog.SelectedHotkeyInputKind != _hotkeyInputKind
                || dialog.SelectedHotkeyKey != _hotkeyKey
                || dialog.SelectedHotkeyMouseButton != _hotkeyMouseButton;

            if (hotkeyChanged
                && !TryApplyHotkey(
                    dialog.SelectedHotkeyModifiers,
                    dialog.SelectedHotkeyKey,
                    dialog.SelectedHotkeyInputKind,
                    dialog.SelectedHotkeyMouseButton))
            {
                continue;
            }

            if (dialog.AutoStartEnabled != _isAutoStartEnabled)
            {
                ApplyAndPersistAutoStartSetting(dialog.AutoStartEnabled);
            }
            else if (!hotkeyChanged)
            {
                SaveSettings();
            }

            return;
        }
    }

    private void InitializeTrayIcon()
    {
        if (_trayIcon is not null)
        {
            return;
        }

        var trayMenu = new Forms.ContextMenuStrip();
        trayMenu.Items.Add("Pokaz / ukryj", null, (_, _) => ToggleWindowVisibility());
        trayMenu.Items.Add("Zamknij", null, (_, _) => ExitApplication());

        _applicationTrayIcon = LoadApplicationTrayIcon();

        _trayIcon = new Forms.NotifyIcon
        {
            Icon = _applicationTrayIcon ?? Drawing.SystemIcons.Application,
            Text = "AMTool",
            Visible = true,
            ContextMenuStrip = trayMenu
        };

        _trayIcon.DoubleClick += (_, _) => ToggleWindowVisibility();
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WM_HOTKEY && wParam.ToInt32() == HotkeyIdToggleWindow)
        {
            ToggleWindowVisibility();
            handled = true;
        }

        return IntPtr.Zero;
    }

    private async void ToggleWindowVisibility()
    {
        if (_isVisibilityAnimationRunning)
        {
            return;
        }

        if (IsVisible)
        {
            await HideToTrayAsync();
            return;
        }

        await ShowFromTrayAsync();
    }

    public async Task ActivateFromExternalRequestAsync()
    {
        if (_isExitRequested)
        {
            return;
        }

        await WaitForVisibilityAnimationAsync();

        if (!IsLoaded)
        {
            return;
        }

        if (!IsVisible)
        {
            await ShowFromTrayAsync();
            return;
        }

        if (WindowState == WindowState.Minimized)
        {
            WindowState = WindowState.Normal;
        }

        PositionWindow();
        BringWindowToFront();
    }

    private async Task ShowFromTrayAsync()
    {
        _isVisibilityAnimationRunning = true;
        PositionWindow();
        ExitHiddenIdleMode();

        Show();
        UpdateMouseHookState();
        WindowState = WindowState.Normal;
        RefreshShortcutStrip(animateFromCenter: true);
        BringWindowToFront();
        await Task.Delay(GetOrbitAnimationTotalDurationMs(ShortcutOrbitLayer.Children.Count, OrbitEntryDurationMs));
        _isVisibilityAnimationRunning = false;
    }

    private async Task HideToTrayAsync()
    {
        if (!IsVisible || _isVisibilityAnimationRunning)
        {
            return;
        }

        _isVisibilityAnimationRunning = true;
        AnimateOrbitCollapse();
        await Task.Delay(GetOrbitAnimationTotalDurationMs(ShortcutOrbitLayer.Children.Count, OrbitExitDurationMs));
        Hide();
        EnterHiddenIdleMode();
        UpdateMouseHookState();
        _isVisibilityAnimationRunning = false;
    }

    private void HideToTrayImmediate()
    {
        if (!IsVisible)
        {
            return;
        }

        Hide();
        EnterHiddenIdleMode();
        UpdateMouseHookState();
    }

    private void PositionWindow()
    {
        Rect workArea = SystemParameters.WorkArea;

        Left = workArea.Left + (workArea.Width - Width) / 2;
        Top = workArea.Bottom - Height - BottomMargin;
    }

    private async Task WaitForVisibilityAnimationAsync()
    {
        int attempt = 0;

        while (_isVisibilityAnimationRunning && attempt < 20)
        {
            await Task.Delay(50);
            attempt++;
        }
    }

    private void BringWindowToFront()
    {
        Activate();
        Focus();

        Topmost = true;
        Topmost = false;
    }

    private static bool IsForegroundWindowOwnedByCurrentProcess()
    {
        IntPtr foregroundWindow = GetForegroundWindow();

        if (foregroundWindow == IntPtr.Zero)
        {
            return false;
        }

        _ = GetWindowThreadProcessId(foregroundWindow, out uint processId);
        return processId == (uint)Environment.ProcessId;
    }

    private void ExitApplication()
    {
        _isExitRequested = true;
        System.Windows.Application.Current.Shutdown();
    }

    private void EnterHiddenIdleMode()
    {
        if (_isHiddenIdleModeActive)
        {
            return;
        }

        ClearDragState();
        ShortcutOrbitLayer.Children.Clear();
        CenterSlotNameText.Text = string.Empty;
        CenterSlotNameText.Visibility = Visibility.Collapsed;
        _isHiddenIdleModeActive = true;
        ScheduleHiddenIdleCleanup();
    }

    private void ExitHiddenIdleMode()
    {
        CancelHiddenIdleCleanup();

        if (!_isHiddenIdleModeActive)
        {
            return;
        }

        _isHiddenIdleModeActive = false;
    }

    private void ScheduleHiddenIdleCleanup()
    {
        CancelHiddenIdleCleanup();
        _hiddenIdleCleanupCancellation = new CancellationTokenSource();
        _ = PerformDeferredHiddenIdleCleanupAsync(_hiddenIdleCleanupCancellation.Token);
    }

    private void CancelHiddenIdleCleanup()
    {
        if (_hiddenIdleCleanupCancellation is null)
        {
            return;
        }

        try
        {
            _hiddenIdleCleanupCancellation.Cancel();
        }
        catch
        {
        }
        finally
        {
            _hiddenIdleCleanupCancellation.Dispose();
            _hiddenIdleCleanupCancellation = null;
        }
    }

    private async Task PerformDeferredHiddenIdleCleanupAsync(CancellationToken cancellationToken)
    {
        try
        {
            await Task.Delay(HiddenIdleCleanupDelayMs, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            return;
        }

        if (cancellationToken.IsCancellationRequested || !_isHiddenIdleModeActive || IsVisible)
        {
            return;
        }

        _iconCache.Clear();
        TrimProcessMemory();
    }

    private static void TrimProcessMemory()
    {
        try
        {
            System.Runtime.GCSettings.LargeObjectHeapCompactionMode = System.Runtime.GCLargeObjectHeapCompactionMode.CompactOnce;
            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);
            GC.WaitForPendingFinalizers();
            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);

            using Process currentProcess = Process.GetCurrentProcess();
            _ = EmptyWorkingSet(currentProcess.Handle);
        }
        catch
        {
        }
    }

    private static Drawing.Icon? LoadApplicationTrayIcon()
    {
        try
        {
            string executablePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName ?? string.Empty;

            if (string.IsNullOrWhiteSpace(executablePath) || !File.Exists(executablePath))
            {
                return null;
            }

            using Drawing.Icon? sourceIcon = Drawing.Icon.ExtractAssociatedIcon(executablePath);
            return sourceIcon is null ? null : (Drawing.Icon)sourceIcon.Clone();
        }
        catch
        {
            return null;
        }
    }

    private async Task ShowFirstRunTutorialAsync()
    {
        FirstRunTutorialDialog dialog;

        try
        {
            dialog = new FirstRunTutorialDialog();
        }
        catch
        {
            System.Windows.MessageBox.Show(
                "Nie udalo sie przygotowac tutorialu pierwszego uruchomienia.",
                "AMTool",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            await ShowFromTrayAsync();
            return;
        }

        if (TryShowDialog(dialog, "Nie udalo sie otworzyc tutorialu pierwszego uruchomienia.", out bool? dialogResult)
            && dialogResult == true)
        {
            _hasCompletedTutorial = true;
            SaveSettings();
        }

        await ShowFromTrayAsync();
    }

    private void LoadShortcuts()
    {
        _shortcuts.Clear();

        if (!File.Exists(_storagePath))
        {
            return;
        }

        try
        {
            string json = File.ReadAllText(_storagePath);
            List<AppShortcutEntry>? loaded = JsonSerializer.Deserialize<List<AppShortcutEntry>>(json);

            if (loaded is null)
            {
                return;
            }

            foreach (AppShortcutEntry entry in loaded)
            {
                if (!string.IsNullOrWhiteSpace(entry.ShortcutPath) && File.Exists(entry.ShortcutPath))
                {
                    _shortcuts.Add(entry);
                }
            }
        }
        catch
        {
        }
    }

    private void SaveShortcuts()
    {
        string? directory = Path.GetDirectoryName(_storagePath);

        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        string json = JsonSerializer.Serialize(_shortcuts, JsonOptions);
        File.WriteAllText(_storagePath, json);
    }

    private void RefreshShortcutStrip(bool animateFromCenter = false)
    {
        ClampShortcutScrollIndex();
        ShortcutOrbitLayer.Children.Clear();

        List<object?> visibleOrbitItems = GetVisibleOrbitItems(_shortcutScrollIndex);
        System.Windows.Point collapsePoint = GetOrbitCollapsePoint();
        UpdateCenterSlotName(visibleOrbitItems);

        for (int slotIndex = 0; slotIndex < visibleOrbitItems.Count; slotIndex++)
        {
            object? orbitItem = visibleOrbitItems[slotIndex];
            Border tile = CreateTileForOrbitItem(orbitItem);

            System.Windows.Point position = OrbitSlotPositions[slotIndex];

            PrepareTileForAnimation(tile);

            if (animateFromCenter)
            {
                SetTileState(tile, collapsePoint, 0, OrbitEntryScale, GetEntryAngle(slotIndex));
            }
            else
            {
                SetTileState(tile, position, 1, 1, 0);
            }

            ShortcutOrbitLayer.Children.Add(tile);

            if (animateFromCenter)
            {
                AnimateTileToPosition(tile, position, 1, 1, 0, slotIndex, OrbitEntryDurationMs);
            }
        }

        UpdateInfoTooltip(_shortcuts.Count + 1);
    }

    private void UpdateCenterSlotName(List<object?> visibleOrbitItems)
    {
        object? middleOrbitItem = visibleOrbitItems.Count > 2
            ? visibleOrbitItems[2]
            : null;

        if (middleOrbitItem is AppShortcutEntry shortcut)
        {
            CenterSlotNameText.Text = shortcut.DisplayName;
            CenterSlotNameText.Visibility = Visibility.Visible;
            return;
        }

        CenterSlotNameText.Text = string.Empty;
        CenterSlotNameText.Visibility = Visibility.Collapsed;
    }

    private List<object?> BuildOrbitItems()
    {
        var orbitItems = new List<object?>(_shortcuts.Count + 1);

        foreach (AppShortcutEntry shortcut in _shortcuts)
        {
            orbitItems.Add(shortcut);
        }

        orbitItems.Add(null);
        return orbitItems;
    }

    private List<object?> GetVisibleOrbitItems(int scrollIndex)
    {
        List<object?> orbitItems = BuildOrbitItems();
        int visibleCount = Math.Min(MaxVisibleOrbitTiles, orbitItems.Count);
        var visibleOrbitItems = new List<object?>(visibleCount);

        for (int index = 0; index < visibleCount; index++)
        {
            int orbitIndex = (scrollIndex + index) % orbitItems.Count;
            visibleOrbitItems.Add(orbitItems[orbitIndex]);
        }

        return visibleOrbitItems;
    }

    private void ClampShortcutScrollIndex()
    {
        _shortcutScrollIndex = Math.Clamp(
            _shortcutScrollIndex,
            0,
            GetMaxShortcutScrollIndex(_shortcuts.Count + 1));
    }

    private static int GetMaxShortcutScrollIndex(int totalTileCount)
    {
        return totalTileCount > MaxVisibleOrbitTiles
            ? totalTileCount - 1
            : 0;
    }

    private static int GetLoopedScrollIndex(int currentScrollIndex, int scrollDirection, int maxScrollIndex)
    {
        if (maxScrollIndex == 0)
        {
            return 0;
        }

        return scrollDirection > 0
            ? currentScrollIndex >= maxScrollIndex ? 0 : currentScrollIndex + 1
            : currentScrollIndex <= 0 ? maxScrollIndex : currentScrollIndex - 1;
    }

    private Border CreateTileForOrbitItem(object? orbitItem)
    {
        return orbitItem is AppShortcutEntry shortcut
            ? CreateShortcutTile(shortcut)
            : CreateAddTile();
    }

    private static string BuildOrbitItemKey(object? orbitItem)
    {
        return orbitItem is AppShortcutEntry shortcut
            ? shortcut.ShortcutPath
            : AddOrbitItemKey;
    }

    private static System.Windows.Point GetOrbitBeforePoint()
    {
        System.Windows.Point first = OrbitSlotPositions[0];
        System.Windows.Point second = OrbitSlotPositions[1];

        return new(
            first.X + (first.X - second.X),
            first.Y + (first.Y - second.Y));
    }

    private static System.Windows.Point GetOrbitAfterPoint()
    {
        System.Windows.Point last = OrbitSlotPositions[MaxVisibleOrbitTiles - 1];
        System.Windows.Point beforeLast = OrbitSlotPositions[MaxVisibleOrbitTiles - 2];

        return new(
            last.X + (last.X - beforeLast.X),
            last.Y + (last.Y - beforeLast.Y));
    }

    private void AnimateOrbitScroll(int previousScrollIndex, int nextScrollIndex, int scrollDirection, int durationMs)
    {
        List<object?> previousVisibleOrbitItems = GetVisibleOrbitItems(previousScrollIndex);
        List<object?> nextVisibleOrbitItems = GetVisibleOrbitItems(nextScrollIndex);

        var existingTilesByKey = new Dictionary<string, Border>(StringComparer.OrdinalIgnoreCase);

        foreach (UIElement child in ShortcutOrbitLayer.Children)
        {
            if (child is Border tile && !string.IsNullOrWhiteSpace(tile.Uid))
            {
                existingTilesByKey[tile.Uid] = tile;
            }
        }

        System.Windows.Point entryPoint = scrollDirection > 0
            ? GetOrbitAfterPoint()
            : GetOrbitBeforePoint();
        System.Windows.Point exitPoint = scrollDirection > 0
            ? GetOrbitBeforePoint()
            : GetOrbitAfterPoint();

        var nextKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (int slotIndex = 0; slotIndex < nextVisibleOrbitItems.Count; slotIndex++)
        {
            object? orbitItem = nextVisibleOrbitItems[slotIndex];
            string orbitItemKey = BuildOrbitItemKey(orbitItem);
            nextKeys.Add(orbitItemKey);

            System.Windows.Point targetPosition = OrbitSlotPositions[slotIndex];

            if (existingTilesByKey.TryGetValue(orbitItemKey, out Border? existingTile))
            {
                AnimateTileToPosition(existingTile, targetPosition, 1, 1, 0, 0, durationMs);
                continue;
            }

            Border enteringTile = CreateTileForOrbitItem(orbitItem);
            PrepareTileForAnimation(enteringTile);
            SetTileState(enteringTile, entryPoint, 0, 0.86, 0);
            ShortcutOrbitLayer.Children.Add(enteringTile);
            AnimateTileToPosition(enteringTile, targetPosition, 1, 1, 0, 0, durationMs);
        }

        foreach (object? orbitItem in previousVisibleOrbitItems)
        {
            string orbitItemKey = BuildOrbitItemKey(orbitItem);

            if (nextKeys.Contains(orbitItemKey) || !existingTilesByKey.TryGetValue(orbitItemKey, out Border? exitingTile))
            {
                continue;
            }

            AnimateTileToPosition(exitingTile, exitPoint, 0, OrbitExitScale, 0, 0, durationMs);
        }
    }

    private async Task ProcessOrbitScrollAsync()
    {
        if (_isScrollAnimationRunning)
        {
            return;
        }

        _isScrollAnimationRunning = true;

        try
        {
            while (true)
            {
                if (_isVisibilityAnimationRunning || _pressedShortcutTile is not null || !IsVisible)
                {
                    return;
                }

                int totalTileCount = _shortcuts.Count + 1;
                int maxScrollIndex = GetMaxShortcutScrollIndex(totalTileCount);

                if (maxScrollIndex == 0)
                {
                    return;
                }

                int inputVersionAtStepStart = _orbitScrollInputVersion;
                int scrollDirection = _orbitScrollDirection;
                int durationMs = _isOrbitScrollBurstActive ? OrbitScrollFastDurationMs : OrbitScrollDurationMs;
                int nextScrollIndex = GetLoopedScrollIndex(_shortcutScrollIndex, scrollDirection, maxScrollIndex);

                if (nextScrollIndex == _shortcutScrollIndex)
                {
                    return;
                }

                int previousScrollIndex = _shortcutScrollIndex;
                _shortcutScrollIndex = nextScrollIndex;
                AnimateOrbitScroll(previousScrollIndex, nextScrollIndex, scrollDirection, durationMs);
                await Task.Delay(durationMs);
                RefreshShortcutStrip();

                bool receivedNewInput = _orbitScrollInputVersion != inputVersionAtStepStart;

                if (!receivedNewInput)
                {
                    _isOrbitScrollBurstActive = false;
                    return;
                }

                _isOrbitScrollBurstActive =
                    (DateTime.UtcNow - _lastOrbitScrollInputUtc).TotalMilliseconds <= OrbitScrollBurstThresholdMs;
            }
        }
        finally
        {
            _isScrollAnimationRunning = false;
        }
    }

    private int GetOrbitAnimationTotalDurationMs(int tileCount, int animationDurationMs)
    {
        if (tileCount <= 0)
        {
            return animationDurationMs;
        }

        return animationDurationMs + ((tileCount - 1) * OrbitAnimationStaggerMs);
    }

    private System.Windows.Point GetOrbitCollapsePoint()
    {
        return new(
            (ShortcutOrbitLayer.Width - OrbitTileSize) / 2,
            (ShortcutOrbitLayer.Height - OrbitTileSize) / 2);
    }

    private static double GetEntryAngle(int slotIndex)
    {
        return slotIndex switch
        {
            0 => -OrbitEntryAngle,
            1 => -OrbitEntryAngle * 0.55,
            2 => 0,
            3 => OrbitEntryAngle * 0.55,
            _ => OrbitEntryAngle
        };
    }

    private static void PrepareTileForAnimation(Border tile)
    {
        var scaleTransform = new ScaleTransform(1, 1);
        var rotateTransform = new RotateTransform(0);
        var transformGroup = new TransformGroup();
        transformGroup.Children.Add(scaleTransform);
        transformGroup.Children.Add(rotateTransform);

        tile.RenderTransformOrigin = new System.Windows.Point(0.5, 0.5);
        tile.RenderTransform = transformGroup;
    }

    private static ScaleTransform GetTileScaleTransform(Border tile)
    {
        return (ScaleTransform)((TransformGroup)tile.RenderTransform).Children[0];
    }

    private static RotateTransform GetTileRotateTransform(Border tile)
    {
        return (RotateTransform)((TransformGroup)tile.RenderTransform).Children[1];
    }

    private static void SetTileState(Border tile, System.Windows.Point position, double opacity, double scale, double angle)
    {
        Canvas.SetLeft(tile, position.X);
        Canvas.SetTop(tile, position.Y);
        tile.Opacity = opacity;

        ScaleTransform scaleTransform = GetTileScaleTransform(tile);
        scaleTransform.ScaleX = scale;
        scaleTransform.ScaleY = scale;
        GetTileRotateTransform(tile).Angle = angle;
    }

    private void AnimateTileToPosition(
        Border tile,
        System.Windows.Point targetPosition,
        double targetOpacity,
        double targetScale,
        double targetAngle,
        int slotIndex,
        int durationMs)
    {
        TimeSpan delay = TimeSpan.FromMilliseconds(slotIndex * OrbitAnimationStaggerMs);

        tile.BeginAnimation(Canvas.LeftProperty, CreateAnimation(targetPosition.X, durationMs, delay, OrbitEaseOut));
        tile.BeginAnimation(Canvas.TopProperty, CreateAnimation(targetPosition.Y, durationMs, delay, OrbitEaseOut));
        tile.BeginAnimation(OpacityProperty, CreateAnimation(targetOpacity, durationMs, delay, OrbitEaseOut));

        ScaleTransform scaleTransform = GetTileScaleTransform(tile);
        scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, CreateAnimation(targetScale, durationMs, delay, OrbitEaseOut));
        scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, CreateAnimation(targetScale, durationMs, delay, OrbitEaseOut));

        RotateTransform rotateTransform = GetTileRotateTransform(tile);
        rotateTransform.BeginAnimation(RotateTransform.AngleProperty, CreateAnimation(targetAngle, durationMs, delay, OrbitEaseOut));
    }

    private static DoubleAnimation CreateAnimation(double toValue, int durationMs, TimeSpan delay, IEasingFunction easing)
    {
        return new DoubleAnimation
        {
            To = toValue,
            Duration = TimeSpan.FromMilliseconds(durationMs),
            BeginTime = delay,
            EasingFunction = easing
        };
    }

    private static System.Windows.Point GetTilePosition(Border tile)
    {
        return new(Canvas.GetLeft(tile), Canvas.GetTop(tile));
    }

    private void AnimateTileScale(Border tile, double targetScale, int durationMs)
    {
        ScaleTransform scaleTransform = GetTileScaleTransform(tile);

        if (Math.Abs(scaleTransform.ScaleX - targetScale) < 0.001
            && Math.Abs(scaleTransform.ScaleY - targetScale) < 0.001)
        {
            return;
        }

        scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, CreateAnimation(targetScale, durationMs, TimeSpan.Zero, OrbitEaseOut));
        scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, CreateAnimation(targetScale, durationMs, TimeSpan.Zero, OrbitEaseOut));
    }

    private static System.Windows.Controls.ToolTip CreateTileTooltip(string text)
    {
        return new System.Windows.Controls.ToolTip
        {
            Content = text,
            Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom,
            HasDropShadow = false
        };
    }

    private void HandleTileMouseEnter(Border tile)
    {
        if (_pressedShortcutTile is not null
            || _isVisibilityAnimationRunning
            || _isScrollAnimationRunning
            || ReferenceEquals(tile, _dragTargetTile))
        {
            return;
        }

        AnimateTileScale(tile, HoverTileScale, OrbitHoverDurationMs);
    }

    private void HandleTileMouseLeave(Border tile)
    {
        if (_pressedShortcutTile is not null
            || _isVisibilityAnimationRunning
            || _isScrollAnimationRunning
            || ReferenceEquals(tile, _dragTargetTile))
        {
            return;
        }

        AnimateTileScale(tile, 1, OrbitHoverDurationMs);
    }

    private static bool IsDragGesture(System.Windows.Point startPosition, System.Windows.Point currentPosition)
    {
        return Math.Abs(currentPosition.X - startPosition.X) >= SystemParameters.MinimumHorizontalDragDistance
            || Math.Abs(currentPosition.Y - startPosition.Y) >= SystemParameters.MinimumVerticalDragDistance;
    }

    private static bool IsPointerOverTile(Border tile, System.Windows.Point pointerPosition)
    {
        System.Windows.Point tilePosition = GetTilePosition(tile);

        return pointerPosition.X >= tilePosition.X
            && pointerPosition.X <= tilePosition.X + tile.Width
            && pointerPosition.Y >= tilePosition.Y
            && pointerPosition.Y <= tilePosition.Y + tile.Height;
    }

    private Border? FindShortcutTileAt(System.Windows.Point pointerPosition)
    {
        for (int index = 0; index < ShortcutOrbitLayer.Children.Count; index++)
        {
            if (ShortcutOrbitLayer.Children[index] is not Border tile
                || ReferenceEquals(tile, _pressedShortcutTile)
                || tile.Tag is not AppShortcutEntry
                || !IsPointerOverTile(tile, pointerPosition))
            {
                continue;
            }

            return tile;
        }

        return null;
    }

    private void UpdateDragTargetTile(Border? nextTargetTile)
    {
        if (ReferenceEquals(_dragTargetTile, nextTargetTile))
        {
            return;
        }

        if (_dragTargetTile is not null)
        {
            AnimateTileScale(_dragTargetTile, 1, 120);
        }

        _dragTargetTile = nextTargetTile;

        if (_dragTargetTile is not null)
        {
            AnimateTileScale(_dragTargetTile, DragTargetScale, 120);
        }
    }

    private void ClearDragState()
    {
        if (_dragTargetTile is not null)
        {
            AnimateTileScale(_dragTargetTile, 1, 120);
        }

        _pressedShortcutTile?.ReleaseMouseCapture();
        _pressedShortcutTile = null;
        _pressedShortcut = null;
        _dragTargetTile = null;
        _isShortcutDragActive = false;
    }

    private bool TrySwapShortcuts(AppShortcutEntry firstShortcut, AppShortcutEntry secondShortcut)
    {
        int firstIndex = _shortcuts.FindIndex(item =>
            string.Equals(item.ShortcutPath, firstShortcut.ShortcutPath, StringComparison.OrdinalIgnoreCase));
        int secondIndex = _shortcuts.FindIndex(item =>
            string.Equals(item.ShortcutPath, secondShortcut.ShortcutPath, StringComparison.OrdinalIgnoreCase));

        if (firstIndex < 0 || secondIndex < 0 || firstIndex == secondIndex)
        {
            return false;
        }

        (_shortcuts[firstIndex], _shortcuts[secondIndex]) = (_shortcuts[secondIndex], _shortcuts[firstIndex]);
        return true;
    }

    private async Task AnimateShortcutSwapAsync(
        Border draggedTile,
        AppShortcutEntry draggedShortcut,
        System.Windows.Point sourceSlotPosition,
        Border targetTile,
        AppShortcutEntry targetShortcut)
    {
        System.Windows.Point targetSlotPosition = GetTilePosition(targetTile);

        AnimateTileToPosition(draggedTile, targetSlotPosition, 1, 1, 0, 0, OrbitSwapDurationMs);
        AnimateTileToPosition(targetTile, sourceSlotPosition, 1, 1, 0, 0, OrbitSwapDurationMs);
        await Task.Delay(OrbitSwapDurationMs);

        if (TrySwapShortcuts(draggedShortcut, targetShortcut))
        {
            SaveShortcuts();
        }

        RefreshShortcutStrip();
    }

    private async Task AnimateShortcutReturnAsync(Border tile, System.Windows.Point sourceSlotPosition)
    {
        AnimateTileToPosition(tile, sourceSlotPosition, 1, 1, 0, 0, OrbitSwapDurationMs);
        await Task.Delay(OrbitSwapDurationMs);
        RefreshShortcutStrip();
    }

    private void AnimateOrbitCollapse()
    {
        System.Windows.Point collapsePoint = GetOrbitCollapsePoint();

        for (int index = 0; index < ShortcutOrbitLayer.Children.Count; index++)
        {
            if (ShortcutOrbitLayer.Children[index] is not Border tile)
            {
                continue;
            }

            double angle = index < ShortcutOrbitLayer.Children.Count / 2
                ? -OrbitExitAngle
                : OrbitExitAngle;

            TimeSpan delay = TimeSpan.FromMilliseconds(index * OrbitAnimationStaggerMs);

            tile.BeginAnimation(Canvas.LeftProperty, CreateAnimation(collapsePoint.X, OrbitExitDurationMs, delay, OrbitEaseIn));
            tile.BeginAnimation(Canvas.TopProperty, CreateAnimation(collapsePoint.Y, OrbitExitDurationMs, delay, OrbitEaseIn));
            tile.BeginAnimation(OpacityProperty, CreateAnimation(0, OrbitExitDurationMs, delay, OrbitEaseIn));

            ScaleTransform scaleTransform = GetTileScaleTransform(tile);
            scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, CreateAnimation(OrbitExitScale, OrbitExitDurationMs, delay, OrbitEaseIn));
            scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, CreateAnimation(OrbitExitScale, OrbitExitDurationMs, delay, OrbitEaseIn));

            RotateTransform rotateTransform = GetTileRotateTransform(tile);
            rotateTransform.BeginAnimation(RotateTransform.AngleProperty, CreateAnimation(angle, OrbitExitDurationMs, delay, OrbitEaseIn));
        }
    }

    private void UpdateInfoTooltip(int totalTileCount)
    {
        string scrollInfo = totalTileCount > MaxVisibleOrbitTiles
            ? $"Kolko myszy przewija kolejne skroty w petli. Widok {_shortcutScrollIndex + 1} z {GetMaxShortcutScrollIndex(totalTileCount) + 1}."
            : "Do 5 skrotow widac od razu bez przewijania.";

        InfoBadge.ToolTip = string.Join(
            Environment.NewLine,
            $"{GetCurrentHotkeyText()} wywoluje AMTool.",
            "LPM uruchamia aplikacje.",
            "PPM usuwa skrot z listy.",
            "PPM na ( i ) otwiera ustawienia.",
            $"Autostart: {(_isAutoStartEnabled ? "wlaczony" : "wylaczony")}.",
            $"Liczba skrotow: {_shortcuts.Count}.",
            scrollInfo);
    }

    private Border CreateShortcutTile(AppShortcutEntry shortcut)
    {
        var image = new System.Windows.Controls.Image
        {
            Source = GetIconForShortcut(shortcut),
            Width = ShortcutIconSize,
            Height = ShortcutIconSize,
            Stretch = Stretch.Uniform,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };

        var iconShell = new Border
        {
            Width = ShortcutIconShellSize,
            Height = ShortcutIconShellSize,
            CornerRadius = new CornerRadius(ShortcutIconShellSize / 2),
            Background = System.Windows.Media.Brushes.Transparent,
            BorderBrush = System.Windows.Media.Brushes.Transparent,
            BorderThickness = new Thickness(0),
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            Child = image
        };

        var border = new Border
        {
            Width = OrbitTileSize,
            Height = OrbitTileSize,
            CornerRadius = new CornerRadius(OrbitTileCornerRadius),
            Background = ShortcutTileBackgroundBrush,
            BorderBrush = ShortcutTileBorderBrush,
            BorderThickness = new Thickness(1.2),
            Child = iconShell,
            Cursor = System.Windows.Input.Cursors.Hand,
            Uid = BuildOrbitItemKey(shortcut),
            Tag = shortcut,
            CacheMode = OrbitTileBitmapCache,
            ToolTip = CreateTileTooltip(shortcut.DisplayName)
        };

        ToolTipService.SetInitialShowDelay(border, 280);
        ToolTipService.SetShowDuration(border, 4000);

        var contextMenu = new ContextMenu();
        var removeItem = new MenuItem
        {
            Header = "Usun skrot",
            Tag = shortcut
        };
        removeItem.Click += RemoveShortcutMenuItem_Click;
        contextMenu.Items.Add(removeItem);

        border.ContextMenu = contextMenu;
        border.MouseEnter += ShortcutTile_MouseEnter;
        border.MouseLeave += ShortcutTile_MouseLeave;
        border.MouseLeftButtonDown += ShortcutTile_MouseLeftButtonDown;
        border.MouseMove += ShortcutTile_MouseMove;
        border.MouseLeftButtonUp += ShortcutTile_MouseLeftButtonUp;

        return border;
    }

    private Border CreateAddTile()
    {
        var plusText = new TextBlock
        {
            Text = "+",
            FontSize = 36,
            FontWeight = FontWeights.Light,
            Foreground = System.Windows.Media.Brushes.White,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };

        var border = new Border
        {
            Width = OrbitTileSize,
            Height = OrbitTileSize,
            CornerRadius = new CornerRadius(OrbitTileCornerRadius),
            Background = AddTileBackgroundBrush,
            BorderBrush = AddTileBorderBrush,
            BorderThickness = new Thickness(1.2),
            Child = plusText,
            Cursor = System.Windows.Input.Cursors.Hand,
            Uid = AddOrbitItemKey,
            CacheMode = OrbitTileBitmapCache,
            ToolTip = CreateTileTooltip("Dodaj skrot")
        };

        ToolTipService.SetInitialShowDelay(border, 280);
        ToolTipService.SetShowDuration(border, 3000);
        border.MouseEnter += AddTile_MouseEnter;
        border.MouseLeave += AddTile_MouseLeave;
        border.MouseLeftButtonUp += AddShortcutTile_Click;
        return border;
    }

    private async void OrbitScene_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (_isVisibilityAnimationRunning || _pressedShortcutTile is not null)
        {
            return;
        }

        int totalTileCount = _shortcuts.Count + 1;
        int maxScrollIndex = GetMaxShortcutScrollIndex(totalTileCount);

        if (maxScrollIndex == 0)
        {
            return;
        }

        int scrollDirection = e.Delta < 0 ? 1 : -1;
        DateTime now = DateTime.UtcNow;
        bool useFastScroll =
            (now - _lastOrbitScrollInputUtc).TotalMilliseconds <= OrbitScrollBurstThresholdMs;

        _orbitScrollDirection = scrollDirection;
        _orbitScrollInputVersion++;
        _lastOrbitScrollInputUtc = now;
        _isOrbitScrollBurstActive = useFastScroll;

        await ProcessOrbitScrollAsync();
        e.Handled = true;
    }

    private void ShortcutTile_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (_isVisibilityAnimationRunning
            || _isScrollAnimationRunning
            || sender is not Border border
            || border.Tag is not AppShortcutEntry shortcut)
        {
            return;
        }

        _pressedShortcutTile = border;
        _pressedShortcut = shortcut;
        _dragStartPointerPosition = e.GetPosition(ShortcutOrbitLayer);
        _dragSourceSlotPosition = GetTilePosition(border);
        _isShortcutDragActive = false;
        UpdateDragTargetTile(null);

        System.Windows.Controls.Panel.SetZIndex(border, 1000);
        border.CaptureMouse();
        e.Handled = true;
    }

    private void ShortcutTile_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (_pressedShortcutTile is null || _pressedShortcut is null || e.LeftButton != MouseButtonState.Pressed)
        {
            return;
        }

        System.Windows.Point pointerPosition = e.GetPosition(ShortcutOrbitLayer);

        if (!_isShortcutDragActive)
        {
            if (!IsDragGesture(_dragStartPointerPosition, pointerPosition))
            {
                return;
            }

            _isShortcutDragActive = true;
            AnimateTileScale(_pressedShortcutTile, DraggedTileScale, 120);
            _pressedShortcutTile.Opacity = 0.92;
        }

        Canvas.SetLeft(_pressedShortcutTile, pointerPosition.X - (OrbitTileSize / 2));
        Canvas.SetTop(_pressedShortcutTile, pointerPosition.Y - (OrbitTileSize / 2));
        UpdateDragTargetTile(FindShortcutTileAt(pointerPosition));
        e.Handled = true;
    }

    private async void ShortcutTile_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (_pressedShortcutTile is null || _pressedShortcut is null)
        {
            return;
        }

        Border draggedTile = _pressedShortcutTile;
        AppShortcutEntry draggedShortcut = _pressedShortcut;
        Border? targetTile = _dragTargetTile;
        System.Windows.Point sourceSlotPosition = _dragSourceSlotPosition;
        bool wasDragging = _isShortcutDragActive;

        ClearDragState();

        if (!wasDragging)
        {
            LaunchShortcut(draggedShortcut);
            e.Handled = true;
            return;
        }

        draggedTile.Opacity = 1;
        System.Windows.Controls.Panel.SetZIndex(draggedTile, 0);

        if (targetTile is not null && targetTile.Tag is AppShortcutEntry targetShortcut)
        {
            await AnimateShortcutSwapAsync(draggedTile, draggedShortcut, sourceSlotPosition, targetTile, targetShortcut);
        }
        else
        {
            await AnimateShortcutReturnAsync(draggedTile, sourceSlotPosition);
        }

        e.Handled = true;
    }

    private void ShortcutTile_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (sender is Border border)
        {
            HandleTileMouseEnter(border);
        }
    }

    private void ShortcutTile_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (sender is Border border)
        {
            HandleTileMouseLeave(border);
        }
    }

    private void AddTile_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (sender is Border border)
        {
            HandleTileMouseEnter(border);
        }
    }

    private void AddTile_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (sender is Border border)
        {
            HandleTileMouseLeave(border);
        }
    }

    private void InfoBadge_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
    {
        OpenSettingsDialog();
        e.Handled = true;
    }

    private void AddShortcutTile_Click(object sender, MouseButtonEventArgs e)
    {
        ShowAddShortcutDialog();
    }

    private void RemoveShortcutMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem menuItem && menuItem.Tag is AppShortcutEntry shortcut)
        {
            RemoveShortcut(shortcut);
        }
    }

    private void ShowAddShortcutDialog()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "Wybierz aplikacje lub skrot",
            Filter = "Aplikacje i skroty (*.exe;*.lnk)|*.exe;*.lnk|Wszystkie pliki (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false
        };

        if (!TryShowFileDialog(dialog, out bool? dialogResult) || dialogResult != true)
        {
            return;
        }

        try
        {
            AppShortcutEntry shortcut = CreateShortcutEntry(dialog.FileName);
            bool alreadyExists = _shortcuts.Exists(item =>
                string.Equals(item.ShortcutPath, shortcut.ShortcutPath, StringComparison.OrdinalIgnoreCase));

            if (alreadyExists)
            {
                return;
            }

            _shortcuts.Add(shortcut);
            SaveShortcuts();
            RefreshShortcutStrip();
        }
        catch
        {
            System.Windows.MessageBox.Show(
                "Nie udalo sie dodac wybranej aplikacji.",
                "AMTool",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void LaunchShortcut(AppShortcutEntry shortcut)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = shortcut.ShortcutPath,
                UseShellExecute = true,
                WorkingDirectory = Path.GetDirectoryName(shortcut.ShortcutPath) ?? string.Empty
            });

            _ = HideToTrayAsync();
        }
        catch
        {
            System.Windows.MessageBox.Show(
                "Nie udalo sie uruchomic aplikacji.",
                "AMTool",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void RemoveShortcut(AppShortcutEntry shortcut)
    {
        int index = _shortcuts.FindIndex(item =>
            string.Equals(item.ShortcutPath, shortcut.ShortcutPath, StringComparison.OrdinalIgnoreCase));

        if (index < 0)
        {
            return;
        }

        _shortcuts.RemoveAt(index);
        SaveShortcuts();
        RefreshShortcutStrip();
    }

    private AppShortcutEntry CreateShortcutEntry(string shortcutPath)
    {
        string fullPath = Path.GetFullPath(shortcutPath);
        string extension = Path.GetExtension(fullPath);
        string? targetPath = null;
        string description = string.Empty;

        if (string.Equals(extension, ".lnk", StringComparison.OrdinalIgnoreCase))
        {
            (targetPath, description) = ResolveShortcut(fullPath);
        }

        string resolvedPath = !string.IsNullOrWhiteSpace(targetPath) && File.Exists(targetPath)
            ? targetPath
            : fullPath;

        string displayName = BuildDisplayName(fullPath, resolvedPath);
        string resolvedDescription = BuildDescription(fullPath, resolvedPath, description);

        return new AppShortcutEntry
        {
            DisplayName = displayName,
            Description = resolvedDescription,
            ShortcutPath = fullPath,
            TargetPath = targetPath
        };
    }

    private static string BuildDisplayName(string shortcutPath, string resolvedPath)
    {
        if (string.Equals(Path.GetExtension(shortcutPath), ".lnk", StringComparison.OrdinalIgnoreCase))
        {
            return Path.GetFileNameWithoutExtension(shortcutPath);
        }

        try
        {
            FileVersionInfo version = FileVersionInfo.GetVersionInfo(resolvedPath);

            if (!string.IsNullOrWhiteSpace(version.ProductName))
            {
                return version.ProductName;
            }

            if (!string.IsNullOrWhiteSpace(version.FileDescription))
            {
                return version.FileDescription;
            }
        }
        catch
        {
        }

        return Path.GetFileNameWithoutExtension(resolvedPath);
    }

    private static string BuildDescription(string shortcutPath, string resolvedPath, string description)
    {
        if (!string.IsNullOrWhiteSpace(description))
        {
            return description;
        }

        try
        {
            FileVersionInfo version = FileVersionInfo.GetVersionInfo(resolvedPath);

            if (!string.IsNullOrWhiteSpace(version.FileDescription))
            {
                return version.FileDescription;
            }
        }
        catch
        {
        }

        return string.Equals(Path.GetExtension(shortcutPath), ".lnk", StringComparison.OrdinalIgnoreCase)
            ? resolvedPath
            : shortcutPath;
    }

    private static (string? TargetPath, string Description) ResolveShortcut(string shortcutPath)
    {
        object? shell = null;
        object? shortcut = null;

        try
        {
            Type? shellType = Type.GetTypeFromProgID("WScript.Shell");

            if (shellType is null)
            {
                return (null, string.Empty);
            }

            shell = Activator.CreateInstance(shellType);

            if (shell is null)
            {
                return (null, string.Empty);
            }

            shortcut = shellType.InvokeMember(
                "CreateShortcut",
                BindingFlags.InvokeMethod,
                null,
                shell,
                [shortcutPath]);

            if (shortcut is null)
            {
                return (null, string.Empty);
            }

            Type shortcutType = shortcut.GetType();
            string? targetPath = shortcutType.InvokeMember("TargetPath", BindingFlags.GetProperty, null, shortcut, null) as string;
            string? description = shortcutType.InvokeMember("Description", BindingFlags.GetProperty, null, shortcut, null) as string;

            return (targetPath, description ?? string.Empty);
        }
        catch
        {
            return (null, string.Empty);
        }
        finally
        {
            if (shortcut is not null && Marshal.IsComObject(shortcut))
            {
                Marshal.ReleaseComObject(shortcut);
            }

            if (shell is not null && Marshal.IsComObject(shell))
            {
                Marshal.ReleaseComObject(shell);
            }
        }
    }

    private ImageSource? GetIconForShortcut(AppShortcutEntry shortcut)
    {
        string iconPath = !string.IsNullOrWhiteSpace(shortcut.TargetPath) && File.Exists(shortcut.TargetPath)
            ? shortcut.TargetPath
            : shortcut.ShortcutPath;

        if (_iconCache.TryGetValue(iconPath, out ImageSource? cached))
        {
            return cached;
        }

        ImageSource? icon = CreateIconFromPath(iconPath) ?? GetDefaultIcon();
        _iconCache[iconPath] = icon;
        return icon;
    }

    private ImageSource? GetDefaultIcon()
    {
        const string key = "__default__";

        if (_iconCache.TryGetValue(key, out ImageSource? cached))
        {
            return cached;
        }

        ImageSource? icon = CreateBitmapSource(Drawing.SystemIcons.Application);
        _iconCache[key] = icon;
        return icon;
    }

    private static ImageSource? CreateIconFromPath(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return null;
            }

            using Drawing.Icon? icon = Drawing.Icon.ExtractAssociatedIcon(path);
            return icon is null ? null : CreateBitmapSource(icon);
        }
        catch
        {
            return null;
        }
    }

    private static ImageSource? CreateBitmapSource(Drawing.Icon icon)
    {
        BitmapSource image = Imaging.CreateBitmapSourceFromHIcon(
            icon.Handle,
            Int32Rect.Empty,
            ShortcutIconBitmapSizeOptions);

        image.Freeze();
        return image;
    }

    private static System.Windows.Media.Brush CreateFrozenBrush(byte alpha, byte red, byte green, byte blue)
    {
        var brush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(alpha, red, green, blue));
        brush.Freeze();
        return brush;
    }

    private static BitmapCache CreateFrozenBitmapCache()
    {
        var bitmapCache = new BitmapCache();
        bitmapCache.Freeze();
        return bitmapCache;
    }

    private static IEasingFunction CreateFrozenCubicEase(EasingMode easingMode)
    {
        var easing = new CubicEase
        {
            EasingMode = easingMode
        };
        easing.Freeze();
        return easing;
    }

    private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MSLLHOOKSTRUCT
    {
        public POINT pt;
        public uint mouseData;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int vKey);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);

    [DllImport("psapi.dll", SetLastError = true)]
    private static extern bool EmptyWorkingSet(IntPtr hProcess);
}
