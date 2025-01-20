using System;
using System.ComponentModel;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Markup;
using System.Windows;

namespace Rubberduck.UI.UnitTesting
{
    [MarkupExtensionReturnType(typeof(object))]
    public class LazyBindingExtension : MarkupExtension
    {
        private Binding binding;
        private UIElement bindingTarget;
        private DependencyProperty bindingTargetProperty;

        public LazyBindingExtension()
        {
        }

        public LazyBindingExtension(PropertyPath path) : this()
        {
            Path = path;
        }

        public IValueConverter Converter { get; set; }
        [TypeConverter(typeof(CultureInfoIetfLanguageTagConverter))]
        public CultureInfo ConverterCulture { get; set; }
        public object ConverterParameter { get; set; }
        public string ElementName { get; set; }
        [ConstructorArgument("path")]
        public PropertyPath Path { get; set; }
        public RelativeSource RelativeSource { get; set; }
        public object Source { get; set; }
        public UpdateSourceTrigger UpdateSourceTrigger { get; set; }
        public bool ValidatesOnDataErrors { get; set; }
        public bool ValidatesOnExceptions { get; set; }
        public bool ValidatesOnNotifyDataErrors { get; set; }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            var valueProvider = serviceProvider.GetService(typeof(IProvideValueTarget)) as IProvideValueTarget;
            if (valueProvider == null)
                return null;

            bindingTarget = valueProvider.TargetObject as UIElement;
            if (bindingTarget == null)
                throw new NotSupportedException($"The target must be a UIElement, '{valueProvider.TargetObject}' is not valid.");

            bindingTargetProperty = valueProvider.TargetProperty as DependencyProperty;
            if (bindingTargetProperty == null)
                throw new NotSupportedException($"The target property must be a DependencyProperty, '{valueProvider.TargetProperty}' is not valid.");

            InitializeBinding();
            SetVisibilityHandler();

            return this;
        }

        private void InitializeBinding()
        {
            binding = new Binding
            {
                Path = Path,
                Converter = Converter,
                ConverterCulture = ConverterCulture,
                ConverterParameter = ConverterParameter
            };

            if (!string.IsNullOrEmpty(ElementName))
                binding.ElementName = ElementName;

            if (RelativeSource != null)
                binding.RelativeSource = RelativeSource;

            if (Source != null)
                binding.Source = Source;

            binding.UpdateSourceTrigger = UpdateSourceTrigger;
            binding.ValidatesOnDataErrors = ValidatesOnDataErrors;
            binding.ValidatesOnExceptions = ValidatesOnExceptions;
            binding.ValidatesOnNotifyDataErrors = ValidatesOnNotifyDataErrors;
        }

        private void SetVisibilityHandler()
        {
            bindingTarget.IsVisibleChanged += OnIsVisibleChanged;
        }

        private void OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            UpdateBinding();
        }

        private void UpdateBinding()
        {
            if (bindingTarget.IsVisible && !IsBindingActive())
                ApplyBinding();
            else if (!bindingTarget.IsVisible)
                ClearBinding();
        }

        private bool IsBindingActive()
        {
            return BindingOperations.GetBinding(bindingTarget, bindingTargetProperty) != null;
        }

        private void ApplyBinding()
        {
            if (!IsBindingActive())
                BindingOperations.SetBinding(bindingTarget, bindingTargetProperty, binding);
        }

        private void ClearBinding()
        {
            if (IsBindingActive())
                BindingOperations.ClearBinding(bindingTarget, bindingTargetProperty);
        }
    }
}
