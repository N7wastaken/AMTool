using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
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
using DrawingImaging = System.Drawing.Imaging;
using Forms = System.Windows.Forms;

namespace AMTool;

public partial class MainWindow : Window
{
    private const int WM_HOTKEY = 0x0312;
    private const int HotkeyIdToggleWindow = 9001;
    private const uint MOD_CONTROL = 0x0002;
    private const uint MOD_SHIFT = 0x0004;
    private const uint MOD_ALT = 0x0001;
    private const uint MOD_WIN = 0x0008;
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
    private const int OrbitScrollDurationMs = 220;
    private const int OrbitSwapDurationMs = 240;
    private const int OrbitAnimationStaggerMs = 34;
    private const int OrbitHoverDurationMs = 140;
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
    private Forms.NotifyIcon? _trayIcon;
    private bool _isExitRequested;
    private bool _isHotkeyRegistered;
    private bool _isVisibilityAnimationRunning;
    private bool _isScrollAnimationRunning;
    private bool _isShortcutDragActive;
    private int _shortcutScrollIndex;
    private bool _isAutoStartEnabled;
    private ModifierKeys _hotkeyModifiers = ModifierKeys.Control | ModifierKeys.Shift;
    private Key _hotkeyKey = Key.Q;
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
        Closing += MainWindow_Closing;
        Closed += MainWindow_Closed;

        LoadSettings();
        LoadShortcuts();
        RefreshShortcutStrip();
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

            if (settings is not null && HotkeyUtilities.IsValidHotkey(settings.HotkeyModifiers, settings.HotkeyKey))
            {
                _hotkeyModifiers = HotkeyUtilities.SanitizeModifiers(settings.HotkeyModifiers);
                _hotkeyKey = settings.HotkeyKey;
            }

            _isAutoStartEnabled = settings?.AutoStartEnabled == true;
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
                    HotkeyKey = _hotkeyKey,
                    AutoStartEnabled = _isAutoStartEnabled
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
        return HotkeyUtilities.FormatHotkey(_hotkeyModifiers, _hotkeyKey);
    }

    private void MainWindow_SourceInitialized(object? sender, EventArgs e)
    {
        IntPtr hwnd = new WindowInteropHelper(this).Handle;

        _hwndSource = HwndSource.FromHwnd(hwnd);
        _hwndSource?.AddHook(WndProc);

        RegisterHotkey(hwnd);
        InitializeTrayIcon();
    }

    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        HideToTrayImmediate();
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

        if (_trayIcon is not null)
        {
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            _trayIcon = null;
        }
    }

    private void RegisterHotkey(IntPtr hwnd)
    {
        _isHotkeyRegistered = RegisterHotKey(
            hwnd,
            HotkeyIdToggleWindow,
            ToNativeHotkeyModifiers(_hotkeyModifiers),
            (uint)KeyInterop.VirtualKeyFromKey(_hotkeyKey));

        if (!_isHotkeyRegistered)
        {
            System.Windows.MessageBox.Show(
                $"Nie udalo sie zarejestrowac skrotu {GetCurrentHotkeyText()}.",
                "AMTool",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private uint ToNativeHotkeyModifiers(ModifierKeys modifiers)
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

        return nativeModifiers;
    }

    private bool TryApplyHotkey(ModifierKeys newModifiers, Key newKey)
    {
        if (_hwndSource is null || !HotkeyUtilities.IsValidHotkey(newModifiers, newKey))
        {
            return false;
        }

        IntPtr hwnd = new WindowInteropHelper(this).Handle;
        ModifierKeys previousModifiers = _hotkeyModifiers;
        Key previousKey = _hotkeyKey;

        if (_isHotkeyRegistered)
        {
            UnregisterHotKey(hwnd, HotkeyIdToggleWindow);
            _isHotkeyRegistered = false;
        }

        _hotkeyModifiers = HotkeyUtilities.SanitizeModifiers(newModifiers);
        _hotkeyKey = newKey;
        _isHotkeyRegistered = RegisterHotKey(
            hwnd,
            HotkeyIdToggleWindow,
            ToNativeHotkeyModifiers(_hotkeyModifiers),
            (uint)KeyInterop.VirtualKeyFromKey(_hotkeyKey));

        if (_isHotkeyRegistered)
        {
            SaveSettings();
            UpdateHotkeyUi();
            return true;
        }

        _hotkeyModifiers = previousModifiers;
        _hotkeyKey = previousKey;
        _isHotkeyRegistered = RegisterHotKey(
            hwnd,
            HotkeyIdToggleWindow,
            ToNativeHotkeyModifiers(_hotkeyModifiers),
            (uint)KeyInterop.VirtualKeyFromKey(_hotkeyKey));

        System.Windows.MessageBox.Show(
            $"Nie udalo sie ustawic skrotu {HotkeyUtilities.FormatHotkey(newModifiers, newKey)}. Ten skrot moze byc juz zajety.",
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
                dialog = new SettingsDialog(_hotkeyModifiers, _hotkeyKey, _isAutoStartEnabled)
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

            bool hotkeyChanged = dialog.SelectedHotkeyModifiers != _hotkeyModifiers || dialog.SelectedHotkeyKey != _hotkeyKey;

            if (hotkeyChanged && !TryApplyHotkey(dialog.SelectedHotkeyModifiers, dialog.SelectedHotkeyKey))
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

        _trayIcon = new Forms.NotifyIcon
        {
            Icon = Drawing.SystemIcons.Application,
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

    private async Task ShowFromTrayAsync()
    {
        _isVisibilityAnimationRunning = true;
        PositionWindow();

        Show();
        WindowState = WindowState.Normal;
        RefreshShortcutStrip(animateFromCenter: true);
        Activate();
        Focus();
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
        _isVisibilityAnimationRunning = false;
    }

    private void HideToTrayImmediate()
    {
        if (!IsVisible)
        {
            return;
        }

        Hide();
    }

    private void PositionWindow()
    {
        Rect workArea = SystemParameters.WorkArea;

        Left = workArea.Left + (workArea.Width - Width) / 2;
        Top = workArea.Bottom - Height - BottomMargin;
    }

    private void ExitApplication()
    {
        _isExitRequested = true;
        System.Windows.Application.Current.Shutdown();
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

    private void AnimateOrbitScroll(int previousScrollIndex, int nextScrollIndex, int scrollDirection)
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
                AnimateTileToPosition(existingTile, targetPosition, 1, 1, 0, 0, OrbitScrollDurationMs);
                continue;
            }

            Border enteringTile = CreateTileForOrbitItem(orbitItem);
            PrepareTileForAnimation(enteringTile);
            SetTileState(enteringTile, entryPoint, 0, 0.86, 0);
            ShortcutOrbitLayer.Children.Add(enteringTile);
            AnimateTileToPosition(enteringTile, targetPosition, 1, 1, 0, 0, OrbitScrollDurationMs);
        }

        foreach (object? orbitItem in previousVisibleOrbitItems)
        {
            string orbitItemKey = BuildOrbitItemKey(orbitItem);

            if (nextKeys.Contains(orbitItemKey) || !existingTilesByKey.TryGetValue(orbitItemKey, out Border? exitingTile))
            {
                continue;
            }

            AnimateTileToPosition(exitingTile, exitPoint, 0, OrbitExitScale, 0, 0, OrbitScrollDurationMs);
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
        var easing = new CubicEase { EasingMode = EasingMode.EaseOut };

        tile.BeginAnimation(Canvas.LeftProperty, CreateAnimation(targetPosition.X, durationMs, delay, easing));
        tile.BeginAnimation(Canvas.TopProperty, CreateAnimation(targetPosition.Y, durationMs, delay, easing));
        tile.BeginAnimation(OpacityProperty, CreateAnimation(targetOpacity, durationMs, delay, easing));

        ScaleTransform scaleTransform = GetTileScaleTransform(tile);
        scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, CreateAnimation(targetScale, durationMs, delay, easing));
        scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, CreateAnimation(targetScale, durationMs, delay, easing));

        RotateTransform rotateTransform = GetTileRotateTransform(tile);
        rotateTransform.BeginAnimation(RotateTransform.AngleProperty, CreateAnimation(targetAngle, durationMs, delay, easing));
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
        var easing = new CubicEase { EasingMode = EasingMode.EaseOut };
        ScaleTransform scaleTransform = GetTileScaleTransform(tile);
        scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, CreateAnimation(targetScale, durationMs, TimeSpan.Zero, easing));
        scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, CreateAnimation(targetScale, durationMs, TimeSpan.Zero, easing));
    }

    private static System.Windows.Controls.ToolTip CreateTileTooltip(string text)
    {
        return new System.Windows.Controls.ToolTip
        {
            Content = text,
            Placement = System.Windows.Controls.Primitives.PlacementMode.Mouse,
            HasDropShadow = true
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
            var easing = new CubicEase { EasingMode = EasingMode.EaseIn };

            tile.BeginAnimation(Canvas.LeftProperty, CreateAnimation(collapsePoint.X, OrbitExitDurationMs, delay, easing));
            tile.BeginAnimation(Canvas.TopProperty, CreateAnimation(collapsePoint.Y, OrbitExitDurationMs, delay, easing));
            tile.BeginAnimation(OpacityProperty, CreateAnimation(0, OrbitExitDurationMs, delay, easing));

            ScaleTransform scaleTransform = GetTileScaleTransform(tile);
            scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, CreateAnimation(OrbitExitScale, OrbitExitDurationMs, delay, easing));
            scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, CreateAnimation(OrbitExitScale, OrbitExitDurationMs, delay, easing));

            RotateTransform rotateTransform = GetTileRotateTransform(tile);
            rotateTransform.BeginAnimation(RotateTransform.AngleProperty, CreateAnimation(angle, OrbitExitDurationMs, delay, easing));
        }
    }

    private void UpdateInfoTooltip(int totalTileCount)
    {
        string scrollInfo = totalTileCount > MaxVisibleOrbitTiles
            ? $"Kolko myszy przewija kolejne skroty w petli. Widok {_shortcutScrollIndex + 1} z {GetMaxShortcutScrollIndex(totalTileCount) + 1}."
            : "Do 5 skrotow widac od razu bez przewijania.";

        InfoBadge.ToolTip = string.Join(
            Environment.NewLine,
            $"{GetCurrentHotkeyText()} pokazuje lub chowa panel.",
            "LPM uruchamia aplikacje.",
            "PPM usuwa skrot z listy.",
            "PPM na i otwiera ustawienia.",
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
            Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(42, 244, 247, 251)),
            BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(92, 244, 247, 251)),
            BorderThickness = new Thickness(1.2),
            Child = iconShell,
            Cursor = System.Windows.Input.Cursors.Hand,
            Uid = BuildOrbitItemKey(shortcut),
            Tag = shortcut,
            ToolTip = CreateTileTooltip(shortcut.DisplayName)
        };

        ToolTipService.SetInitialShowDelay(border, 120);
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
            Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(28, 244, 247, 251)),
            BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(118, 244, 247, 251)),
            BorderThickness = new Thickness(1.2),
            Child = plusText,
            Cursor = System.Windows.Input.Cursors.Hand,
            Uid = AddOrbitItemKey,
            ToolTip = CreateTileTooltip("Dodaj skrot")
        };

        ToolTipService.SetInitialShowDelay(border, 120);
        ToolTipService.SetShowDuration(border, 3000);
        border.MouseEnter += AddTile_MouseEnter;
        border.MouseLeave += AddTile_MouseLeave;
        border.MouseLeftButtonUp += AddShortcutTile_Click;
        return border;
    }

    private async void OrbitScene_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (_isVisibilityAnimationRunning || _isScrollAnimationRunning || _pressedShortcutTile is not null)
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
        int nextScrollIndex = GetLoopedScrollIndex(_shortcutScrollIndex, scrollDirection, maxScrollIndex);

        if (nextScrollIndex == _shortcutScrollIndex)
        {
            return;
        }

        int previousScrollIndex = _shortcutScrollIndex;

        _isScrollAnimationRunning = true;
        _shortcutScrollIndex = nextScrollIndex;
        AnimateOrbitScroll(previousScrollIndex, nextScrollIndex, scrollDirection);
        await Task.Delay(OrbitScrollDurationMs);
        RefreshShortcutStrip();
        _isScrollAnimationRunning = false;
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
        using var bitmap = icon.ToBitmap();
        using var stream = new MemoryStream();

        bitmap.Save(stream, DrawingImaging.ImageFormat.Png);
        stream.Position = 0;

        var image = new BitmapImage();
        image.BeginInit();
        image.CacheOption = BitmapCacheOption.OnLoad;
        image.StreamSource = stream;
        image.EndInit();
        image.Freeze();

        return image;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
}
