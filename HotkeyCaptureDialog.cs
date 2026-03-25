using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace AMTool;

public sealed class HotkeyCaptureDialog : Window
{
    private readonly TextBlock _selectionText;
    private readonly System.Windows.Controls.Button _confirmButton;

    public ModifierKeys SelectedModifiers { get; private set; }

    public Key SelectedKey { get; private set; }

    public HotkeyCaptureDialog(ModifierKeys currentModifiers, Key currentKey)
    {
        SelectedModifiers = currentModifiers;
        SelectedKey = currentKey;

        Title = "Ustaw skrot wywolywania";
        Width = 860;
        Height = 500;
        ResizeMode = ResizeMode.NoResize;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(20, 26, 34));
        Foreground = System.Windows.Media.Brushes.White;
        ShowInTaskbar = false;

        PreviewKeyDown += HotkeyCaptureDialog_PreviewKeyDown;
        Loaded += (_, _) => Keyboard.Focus(this);

        var root = new Grid
        {
            Margin = new Thickness(36)
        };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        root.Children.Add(new TextBlock
        {
            Text = "Nacisnij klawisz albo kombinacje dla panelu.",
            FontSize = 20,
            FontWeight = FontWeights.SemiBold
        });

        var helperText = new TextBlock
        {
            Margin = new Thickness(0, 12, 0, 0),
            Text = "Mozesz ustawic pojedynczy klawisz albo kombinacje. Lepiej unikac skrotow, ktorych stale uzywasz w innych programach.",
            TextWrapping = TextWrapping.Wrap,
            FontSize = 15,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(180, 196, 211))
        };
        Grid.SetRow(helperText, 1);
        root.Children.Add(helperText);

        var previewBorder = new Border
        {
            Margin = new Thickness(0, 34, 0, 0),
            Padding = new Thickness(32, 28, 32, 28),
            MinHeight = 196,
            CornerRadius = new CornerRadius(20),
            Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(38, 244, 247, 251)),
            BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(80, 244, 247, 251)),
            BorderThickness = new Thickness(1)
        };
        Grid.SetRow(previewBorder, 2);
        root.Children.Add(previewBorder);

        _selectionText = new TextBlock
        {
            FontSize = 36,
            FontWeight = FontWeights.Bold,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            TextAlignment = TextAlignment.Center,
            TextWrapping = TextWrapping.Wrap,
            Text = HotkeyUtilities.FormatHotkey(SelectedModifiers, SelectedKey)
        };
        previewBorder.Child = _selectionText;

        var buttonPanel = new StackPanel
        {
            Margin = new Thickness(0, 34, 0, 0),
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Right
        };
        Grid.SetRow(buttonPanel, 3);
        root.Children.Add(buttonPanel);

        var cancelButton = new System.Windows.Controls.Button
        {
            Width = 132,
            Height = 46,
            Margin = new Thickness(0, 0, 12, 0),
            Content = "Anuluj",
            IsCancel = true,
            Style = CreateDialogButtonStyle(isPrimary: false)
        };
        buttonPanel.Children.Add(cancelButton);

        _confirmButton = new System.Windows.Controls.Button
        {
            Width = 148,
            Height = 46,
            Content = "Zapisz",
            IsDefault = true,
            Style = CreateDialogButtonStyle(isPrimary: true)
        };
        _confirmButton.Click += (_, _) => DialogResult = true;
        buttonPanel.Children.Add(_confirmButton);

        Content = root;
        UpdateSelectionState();
    }

    private static Style CreateDialogButtonStyle(bool isPrimary)
    {
        var style = new Style(typeof(System.Windows.Controls.Button));

        style.Setters.Add(new Setter(FontSizeProperty, 14.0));
        style.Setters.Add(new Setter(FontWeightProperty, FontWeights.SemiBold));
        style.Setters.Add(new Setter(ForegroundProperty, isPrimary ? System.Windows.Media.Brushes.White : new SolidColorBrush(System.Windows.Media.Color.FromRgb(222, 232, 240))));
        style.Setters.Add(new Setter(BackgroundProperty, isPrimary
            ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 121, 241))
            : new SolidColorBrush(System.Windows.Media.Color.FromRgb(34, 43, 55))));
        style.Setters.Add(new Setter(BorderBrushProperty, isPrimary
            ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(88, 149, 255))
            : new SolidColorBrush(System.Windows.Media.Color.FromArgb(90, 244, 247, 251))));
        style.Setters.Add(new Setter(BorderThicknessProperty, new Thickness(1)));
        style.Setters.Add(new Setter(CursorProperty, System.Windows.Input.Cursors.Hand));

        var template = new ControlTemplate(typeof(System.Windows.Controls.Button));
        var borderFactory = new FrameworkElementFactory(typeof(Border));
        borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(14));
        borderFactory.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(BackgroundProperty));
        borderFactory.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(BorderBrushProperty));
        borderFactory.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(BorderThicknessProperty));

        var presenterFactory = new FrameworkElementFactory(typeof(ContentPresenter));
        presenterFactory.SetValue(HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center);
        presenterFactory.SetValue(VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
        presenterFactory.SetValue(ContentPresenter.ContentProperty, new TemplateBindingExtension(ContentProperty));
        presenterFactory.SetValue(ContentPresenter.ContentTemplateProperty, new TemplateBindingExtension(ContentTemplateProperty));

        borderFactory.AppendChild(presenterFactory);
        template.VisualTree = borderFactory;
        style.Setters.Add(new Setter(TemplateProperty, template));

        var hoverTrigger = new Trigger { Property = IsMouseOverProperty, Value = true };
        hoverTrigger.Setters.Add(new Setter(BackgroundProperty, isPrimary
            ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(72, 138, 255))
            : new SolidColorBrush(System.Windows.Media.Color.FromRgb(44, 54, 68))));
        hoverTrigger.Setters.Add(new Setter(BorderBrushProperty, isPrimary
            ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 173, 255))
            : new SolidColorBrush(System.Windows.Media.Color.FromArgb(135, 244, 247, 251))));
        style.Triggers.Add(hoverTrigger);

        var pressedTrigger = new Trigger { Property = System.Windows.Controls.Primitives.ButtonBase.IsPressedProperty, Value = true };
        pressedTrigger.Setters.Add(new Setter(BackgroundProperty, isPrimary
            ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(38, 102, 216))
            : new SolidColorBrush(System.Windows.Media.Color.FromRgb(28, 36, 47))));
        style.Triggers.Add(pressedTrigger);

        var disabledTrigger = new Trigger { Property = IsEnabledProperty, Value = false };
        disabledTrigger.Setters.Add(new Setter(OpacityProperty, 0.55));
        style.Triggers.Add(disabledTrigger);

        return style;
    }

    private void HotkeyCaptureDialog_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            DialogResult = false;
            e.Handled = true;
            return;
        }

        Key capturedKey = e.Key == Key.System ? e.SystemKey : e.Key;
        ModifierKeys modifiers = HotkeyUtilities.SanitizeModifiers(Keyboard.Modifiers);

        if (HotkeyUtilities.IsModifierOnlyKey(capturedKey))
        {
            _selectionText.Text = modifiers == ModifierKeys.None
                ? "Nacisnij klawisz lub kombinacje"
                : $"{HotkeyUtilities.FormatHotkey(modifiers, Key.None)}+...";
            _confirmButton.IsEnabled = false;
            e.Handled = true;
            return;
        }

        SelectedModifiers = modifiers;
        SelectedKey = capturedKey;
        UpdateSelectionState();
        e.Handled = true;
    }

    private void UpdateSelectionState()
    {
        _selectionText.Text = HotkeyUtilities.FormatHotkey(SelectedModifiers, SelectedKey);
        _confirmButton.IsEnabled = HotkeyUtilities.IsValidHotkey(SelectedModifiers, SelectedKey);
    }
}
