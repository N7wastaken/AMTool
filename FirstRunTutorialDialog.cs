using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace AMTool;

public sealed class FirstRunTutorialDialog : Window
{
    private readonly TutorialStep[] _steps;
    private readonly StackPanel _indicatorPanel;
    private readonly TextBlock _stepCounterText;
    private readonly TextBlock _titleText;
    private readonly TextBlock _descriptionText;
    private readonly TextBlock _detailsText;
    private readonly System.Windows.Controls.Button _backButton;
    private readonly System.Windows.Controls.Button _nextButton;

    private int _currentStepIndex;

    public FirstRunTutorialDialog()
    {
        _steps =
        [
            new TutorialStep(
                "Dodaj pierwszy skrot",
                "Kliknij kafelek +, wybierz plik .exe albo .lnk i dodaj aplikacje do panelu.",
                "Nowa aplikacja zajmie miejsce plusa, a kafelek + przesunie sie na nastepny wolny slot."),
            new TutorialStep(
                "Uruchamiaj i porzadkuj",
                "LPM uruchamia aplikacje. PPM usuwa skrot. Ikony mozesz lapac i zamieniac miejscami przeciagnieciem.",
                "Po najechaniu slot lekko sie powieksza, a nazwa aplikacji pokazuje sie w dymku."),
            new TutorialStep(
                "Skrot, ustawienia i informacje",
                "Domyslnie Ctrl+Shift+Q wywoluje AMTool, ale w ustawieniach mozesz przypisac tez pojedynczy klawisz, inna kombinacje albo przycisk myszy.",
                "PPM na ( i ) otwiera ustawienia, gdzie zmienisz skrot i autostart, a po najechaniu na ( i ) sprawdzisz najwazniejsze informacje w tooltipie.")
        ];

        Title = "Pierwsze uruchomienie";
        Width = 920;
        Height = 620;
        ResizeMode = ResizeMode.NoResize;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(18, 24, 32));
        Foreground = System.Windows.Media.Brushes.White;
        ShowInTaskbar = false;

        var root = new Grid
        {
            Margin = new Thickness(38)
        };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var headerPanel = new StackPanel();
        root.Children.Add(headerPanel);

        headerPanel.Children.Add(new TextBlock
        {
            Text = "Witaj w AMTool",
            FontSize = 30,
            FontWeight = FontWeights.Bold
        });

        headerPanel.Children.Add(new TextBlock
        {
            Margin = new Thickness(0, 14, 0, 0),
            Text = "Krotki tutorial pokaze Ci najwazniejsze rzeczy. Po ukonczeniu nie bedzie juz wyswietlany automatycznie.",
            FontSize = 16,
            TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(180, 196, 211))
        });

        var progressGrid = new Grid
        {
            Margin = new Thickness(0, 26, 0, 0)
        };
        progressGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        progressGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        Grid.SetRow(progressGrid, 1);
        root.Children.Add(progressGrid);

        _indicatorPanel = new StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };
        progressGrid.Children.Add(_indicatorPanel);

        _stepCounterText = new TextBlock
        {
            FontSize = 15,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(180, 196, 211)),
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };
        Grid.SetColumn(_stepCounterText, 1);
        progressGrid.Children.Add(_stepCounterText);

        var contentCard = new Border
        {
            Margin = new Thickness(0, 24, 0, 0),
            Padding = new Thickness(34, 30, 34, 30),
            CornerRadius = new CornerRadius(24),
            Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(38, 244, 247, 251)),
            BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(82, 244, 247, 251)),
            BorderThickness = new Thickness(1)
        };
        Grid.SetRow(contentCard, 2);
        root.Children.Add(contentCard);

        var contentPanel = new Grid();
        contentPanel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        contentPanel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        contentPanel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        contentCard.Child = contentPanel;

        _titleText = new TextBlock
        {
            FontSize = 34,
            FontWeight = FontWeights.Bold,
            TextWrapping = TextWrapping.Wrap
        };
        contentPanel.Children.Add(_titleText);

        _descriptionText = new TextBlock
        {
            Margin = new Thickness(0, 18, 0, 0),
            FontSize = 19,
            TextWrapping = TextWrapping.Wrap,
            Foreground = System.Windows.Media.Brushes.White
        };
        Grid.SetRow(_descriptionText, 1);
        contentPanel.Children.Add(_descriptionText);

        var detailsBorder = new Border
        {
            Margin = new Thickness(0, 24, 0, 0),
            Padding = new Thickness(22, 18, 22, 18),
            CornerRadius = new CornerRadius(18),
            Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(32, 69, 132, 255)),
            BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(90, 138, 188, 255)),
            BorderThickness = new Thickness(1)
        };
        Grid.SetRow(detailsBorder, 2);
        contentPanel.Children.Add(detailsBorder);

        _detailsText = new TextBlock
        {
            FontSize = 15,
            TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(219, 232, 252))
        };
        detailsBorder.Child = _detailsText;

        var buttonPanel = new Grid
        {
            Margin = new Thickness(0, 30, 0, 0)
        };
        buttonPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        buttonPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        buttonPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        buttonPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        Grid.SetRow(buttonPanel, 3);
        root.Children.Add(buttonPanel);

        var laterButton = new System.Windows.Controls.Button
        {
            Width = 142,
            Height = 50,
            Content = "Pozniej",
            IsCancel = true,
            Style = CreateDialogButtonStyle(isPrimary: false)
        };
        buttonPanel.Children.Add(laterButton);

        _backButton = new System.Windows.Controls.Button
        {
            Width = 142,
            Height = 50,
            Margin = new Thickness(0, 0, 14, 0),
            Content = "Wstecz",
            Style = CreateDialogButtonStyle(isPrimary: false)
        };
        _backButton.Click += BackButton_Click;
        Grid.SetColumn(_backButton, 2);
        buttonPanel.Children.Add(_backButton);

        _nextButton = new System.Windows.Controls.Button
        {
            Width = 158,
            Height = 50,
            Content = "Dalej",
            IsDefault = true,
            Style = CreateDialogButtonStyle(isPrimary: true)
        };
        _nextButton.Click += NextButton_Click;
        Grid.SetColumn(_nextButton, 3);
        buttonPanel.Children.Add(_nextButton);

        Content = root;
        UpdateStep();
    }

    private void BackButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentStepIndex <= 0)
        {
            return;
        }

        _currentStepIndex--;
        UpdateStep();
    }

    private void NextButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentStepIndex >= _steps.Length - 1)
        {
            DialogResult = true;
            return;
        }

        _currentStepIndex++;
        UpdateStep();
    }

    private void UpdateStep()
    {
        TutorialStep step = _steps[_currentStepIndex];

        _stepCounterText.Text = $"Krok {_currentStepIndex + 1}/{_steps.Length}";
        _titleText.Text = step.Title;
        _descriptionText.Text = step.Description;
        _detailsText.Text = step.Details;

        _backButton.IsEnabled = _currentStepIndex > 0;
        _nextButton.Content = _currentStepIndex == _steps.Length - 1 ? "Zaczynam" : "Dalej";

        _indicatorPanel.Children.Clear();

        for (int index = 0; index < _steps.Length; index++)
        {
            bool isActive = index == _currentStepIndex;

            _indicatorPanel.Children.Add(new Border
            {
                Width = isActive ? 34 : 12,
                Height = 12,
                Margin = new Thickness(0, 0, 10, 0),
                CornerRadius = new CornerRadius(6),
                Background = isActive
                    ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(80, 142, 255))
                    : new SolidColorBrush(System.Windows.Media.Color.FromArgb(62, 244, 247, 251))
            });
        }
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

    private sealed class TutorialStep(string title, string description, string details)
    {
        public string Title { get; } = title;

        public string Description { get; } = description;

        public string Details { get; } = details;
    }
}
