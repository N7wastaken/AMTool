using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace AMTool;

public sealed class SettingsDialog : Window
{
    private readonly TextBlock _hotkeyValueText;
    private readonly System.Windows.Controls.CheckBox _autoStartCheckBox;

    public ModifierKeys SelectedHotkeyModifiers { get; private set; }

    public Key SelectedHotkeyKey { get; private set; }

    public bool AutoStartEnabled => _autoStartCheckBox.IsChecked == true;

    public SettingsDialog(ModifierKeys currentModifiers, Key currentKey, bool autoStartEnabled)
    {
        SelectedHotkeyModifiers = currentModifiers;
        SelectedHotkeyKey = currentKey;

        Title = "Ustawienia";
        Width = 900;
        Height = 580;
        ResizeMode = ResizeMode.NoResize;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(20, 26, 34));
        Foreground = System.Windows.Media.Brushes.White;
        ShowInTaskbar = false;

        var root = new Grid
        {
            Margin = new Thickness(38)
        };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        root.Children.Add(new TextBlock
        {
            Text = "Ustawienia panelu",
            FontSize = 26,
            FontWeight = FontWeights.Bold
        });

        var helperText = new TextBlock
        {
            Margin = new Thickness(0, 14, 0, 0),
            Text = "Tutaj zmienisz skrot wywolywania panelu i ustawisz autostart po zalogowaniu do systemu.",
            TextWrapping = TextWrapping.Wrap,
            FontSize = 16,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(180, 196, 211))
        };
        Grid.SetRow(helperText, 1);
        root.Children.Add(helperText);

        var hotkeyCard = new Border
        {
            Margin = new Thickness(0, 30, 0, 0),
            Padding = new Thickness(30, 28, 30, 28),
            CornerRadius = new CornerRadius(22),
            Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(38, 244, 247, 251)),
            BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(80, 244, 247, 251)),
            BorderThickness = new Thickness(1)
        };
        Grid.SetRow(hotkeyCard, 2);
        root.Children.Add(hotkeyCard);

        var hotkeyGrid = new Grid();
        hotkeyGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        hotkeyGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        hotkeyCard.Child = hotkeyGrid;

        var hotkeyTextPanel = new StackPanel();
        hotkeyGrid.Children.Add(hotkeyTextPanel);

        hotkeyTextPanel.Children.Add(new TextBlock
        {
            Text = "Skrot wywolywania",
            FontSize = 16,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(180, 196, 211))
        });

        _hotkeyValueText = new TextBlock
        {
            Margin = new Thickness(0, 12, 0, 0),
            FontSize = 38,
            FontWeight = FontWeights.Bold,
            TextWrapping = TextWrapping.Wrap
        };
        hotkeyTextPanel.Children.Add(_hotkeyValueText);

        var changeHotkeyButton = new System.Windows.Controls.Button
        {
            Width = 164,
            Height = 50,
            Margin = new Thickness(28, 0, 0, 0),
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            Content = "Zmien...",
            Style = CreateDialogButtonStyle(isPrimary: true)
        };
        changeHotkeyButton.Click += ChangeHotkeyButton_Click;
        Grid.SetColumn(changeHotkeyButton, 1);
        hotkeyGrid.Children.Add(changeHotkeyButton);

        var autoStartCard = new Border
        {
            Margin = new Thickness(0, 22, 0, 0),
            Padding = new Thickness(30, 28, 30, 28),
            CornerRadius = new CornerRadius(22),
            Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(26, 244, 247, 251)),
            BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(60, 244, 247, 251)),
            BorderThickness = new Thickness(1)
        };
        Grid.SetRow(autoStartCard, 3);
        root.Children.Add(autoStartCard);

        var autoStartGrid = new Grid();
        autoStartGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        autoStartGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        autoStartCard.Child = autoStartGrid;

        var autoStartTextPanel = new StackPanel
        {
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };
        autoStartGrid.Children.Add(autoStartTextPanel);

        autoStartTextPanel.Children.Add(new TextBlock
        {
            Text = "Autostart",
            FontSize = 21,
            FontWeight = FontWeights.SemiBold,
            Foreground = System.Windows.Media.Brushes.White
        });

        autoStartTextPanel.Children.Add(new TextBlock
        {
            Margin = new Thickness(0, 10, 0, 0),
            Text = "Uruchamiaj AMTool automatycznie po zalogowaniu do systemu.",
            FontSize = 15,
            TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(180, 196, 211))
        });

        _autoStartCheckBox = new System.Windows.Controls.CheckBox
        {
            IsChecked = autoStartEnabled,
            Width = 72,
            Height = 40,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            Cursor = System.Windows.Input.Cursors.Hand,
            Style = CreateToggleCheckBoxStyle()
        };
        Grid.SetColumn(_autoStartCheckBox, 1);
        autoStartGrid.Children.Add(_autoStartCheckBox);

        var buttonPanel = new StackPanel
        {
            Margin = new Thickness(0, 34, 0, 0),
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Right
        };
        Grid.SetRow(buttonPanel, 4);
        root.Children.Add(buttonPanel);

        var cancelButton = new System.Windows.Controls.Button
        {
            Width = 148,
            Height = 50,
            Margin = new Thickness(0, 0, 14, 0),
            Content = "Anuluj",
            IsCancel = true,
            Style = CreateDialogButtonStyle(isPrimary: false)
        };
        buttonPanel.Children.Add(cancelButton);

        var saveButton = new System.Windows.Controls.Button
        {
            Width = 164,
            Height = 50,
            Content = "Zapisz",
            IsDefault = true,
            Style = CreateDialogButtonStyle(isPrimary: true)
        };
        saveButton.Click += (_, _) => DialogResult = true;
        buttonPanel.Children.Add(saveButton);

        Content = root;
        UpdateHotkeyPreview();
    }

    private void ChangeHotkeyButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new HotkeyCaptureDialog(SelectedHotkeyModifiers, SelectedHotkeyKey)
            {
                Owner = this
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            SelectedHotkeyModifiers = dialog.SelectedModifiers;
            SelectedHotkeyKey = dialog.SelectedKey;
            UpdateHotkeyPreview();
        }
        catch
        {
            System.Windows.MessageBox.Show(
                "Nie udalo sie otworzyc okna przypisywania skrotu.",
                "AMTool",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void UpdateHotkeyPreview()
    {
        _hotkeyValueText.Text = HotkeyUtilities.FormatHotkey(SelectedHotkeyModifiers, SelectedHotkeyKey);
    }

    private static Style CreateDialogButtonStyle(bool isPrimary)
    {
        var style = new Style(typeof(System.Windows.Controls.Button));

        style.Setters.Add(new Setter(FontSizeProperty, 15.0));
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

    private static Style CreateToggleCheckBoxStyle()
    {
        var style = new Style(typeof(System.Windows.Controls.CheckBox));

        style.Setters.Add(new Setter(FocusVisualStyleProperty, null));
        style.Setters.Add(new Setter(CursorProperty, System.Windows.Input.Cursors.Hand));

        var template = new ControlTemplate(typeof(System.Windows.Controls.CheckBox));

        var rootFactory = new FrameworkElementFactory(typeof(Grid));

        var trackFactory = new FrameworkElementFactory(typeof(Border));
        trackFactory.Name = "Track";
        trackFactory.SetValue(FrameworkElement.WidthProperty, 72.0);
        trackFactory.SetValue(FrameworkElement.HeightProperty, 40.0);
        trackFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(20));
        trackFactory.SetValue(Border.BackgroundProperty, new SolidColorBrush(System.Windows.Media.Color.FromRgb(44, 54, 68)));
        trackFactory.SetValue(Border.BorderBrushProperty, new SolidColorBrush(System.Windows.Media.Color.FromArgb(105, 244, 247, 251)));
        trackFactory.SetValue(Border.BorderThicknessProperty, new Thickness(1.2));
        rootFactory.AppendChild(trackFactory);

        var thumbFactory = new FrameworkElementFactory(typeof(Border));
        thumbFactory.Name = "Thumb";
        thumbFactory.SetValue(FrameworkElement.WidthProperty, 30.0);
        thumbFactory.SetValue(FrameworkElement.HeightProperty, 30.0);
        thumbFactory.SetValue(FrameworkElement.MarginProperty, new Thickness(5, 0, 0, 0));
        thumbFactory.SetValue(FrameworkElement.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Left);
        thumbFactory.SetValue(FrameworkElement.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
        thumbFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(15));
        thumbFactory.SetValue(Border.BackgroundProperty, System.Windows.Media.Brushes.White);
        thumbFactory.SetValue(Border.BorderBrushProperty, new SolidColorBrush(System.Windows.Media.Color.FromArgb(30, 9, 14, 22)));
        thumbFactory.SetValue(Border.BorderThicknessProperty, new Thickness(1));
        rootFactory.AppendChild(thumbFactory);

        template.VisualTree = rootFactory;

        var hoverTrigger = new Trigger { Property = IsMouseOverProperty, Value = true };
        hoverTrigger.Setters.Add(new Setter(Border.BorderBrushProperty, new SolidColorBrush(System.Windows.Media.Color.FromRgb(132, 186, 255)), "Track"));
        hoverTrigger.Setters.Add(new Setter(Border.BackgroundProperty, new SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 64, 80)), "Track"));
        template.Triggers.Add(hoverTrigger);

        var checkedTrigger = new Trigger { Property = System.Windows.Controls.Primitives.ToggleButton.IsCheckedProperty, Value = true };
        checkedTrigger.Setters.Add(new Setter(Border.BackgroundProperty, new SolidColorBrush(System.Windows.Media.Color.FromRgb(54, 123, 241)), "Track"));
        checkedTrigger.Setters.Add(new Setter(Border.BorderBrushProperty, new SolidColorBrush(System.Windows.Media.Color.FromRgb(129, 179, 255)), "Track"));
        checkedTrigger.Setters.Add(new Setter(FrameworkElement.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Right, "Thumb"));
        checkedTrigger.Setters.Add(new Setter(FrameworkElement.MarginProperty, new Thickness(0, 0, 5, 0), "Thumb"));
        template.Triggers.Add(checkedTrigger);

        var disabledTrigger = new Trigger { Property = IsEnabledProperty, Value = false };
        disabledTrigger.Setters.Add(new Setter(OpacityProperty, 0.55));
        template.Triggers.Add(disabledTrigger);

        style.Setters.Add(new Setter(TemplateProperty, template));
        return style;
    }
}
