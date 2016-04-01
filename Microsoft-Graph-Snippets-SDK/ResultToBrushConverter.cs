//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.UI;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Media;

namespace Microsoft_Graph_Snippets_SDK
{
    /// <summary>
    /// Value converter that translates true to <see cref="Visibility.Visible"/> and false to
    /// <see cref="Visibility.Collapsed"/>.
    /// </summary>
    public sealed class ResultToBrushConverter : IValueConverter
    {
        private static readonly SolidColorBrush NOT_STARTED_BRUSH = new SolidColorBrush(Colors.LightSlateGray);
        private static readonly SolidColorBrush SUCCESS_BRUSH = new SolidColorBrush(Colors.Green);
        private static readonly SolidColorBrush FAILED_BRUSH = new SolidColorBrush(Colors.Red);
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            bool? result = (bool?)value;
            if (result.HasValue)
            {
                return (result.Value) ? SUCCESS_BRUSH : FAILED_BRUSH;
            }
            else
            {
                return NOT_STARTED_BRUSH;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            throw new NotImplementedException();
        }
    }
}


