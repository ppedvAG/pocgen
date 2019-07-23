using ppedv.pocgen.UI.WPF.ViewModels;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace pocgen.Controls
{
    /// <summary>
    /// Interaction logic for PreviewControls.xaml
    /// </summary>
    public partial class PreviewControl : GroupBox, INotifyPropertyChanged
    {
        public PreviewControl()
        {
            InitializeComponent();
            this.DataContext = this;
            MinimumSlides = 0;
            tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "pocgen");
        }
        private string tempPath;

        public int MinimumSlides;
        private int currentSlide;
        public int CurrentSlide
        {
            get => currentSlide;
            set
            {
                Set(ref currentSlide, value);
                ImageSource = includedImages[currentSlide];
            }
        }
        private int maximumSlides;
        public int MaximumSlides
        {
            get => maximumSlides;
            set => Set(ref maximumSlides, value);
        }

        private string imageSource;
        public string ImageSource
        {
            get => imageSource;
            set => Set(ref imageSource,value);
        }

        //private IEnumerable<PowerPointPresentationItem> itemSource;
        //public IEnumerable<PowerPointPresentationItem> ItemSource
        //{
        //    get => itemSource;
        //    set
        //    {
        //        Set(ref itemSource, value);
        //        RefreshUI();
        //    }
        //}



        public IEnumerable<PowerPointPresentationItem> ItemSource
        {
            get => (IEnumerable<PowerPointPresentationItem>)GetValue(ItemSourceProperty); 
            set
            {
                SetValue(ItemSourceProperty, value);
                RefreshUI();
            }
        }

        public static readonly DependencyProperty ItemSourceProperty =
            DependencyProperty.Register(nameof(ItemSource), typeof(IEnumerable<PowerPointPresentationItem>), typeof(PreviewControl), new PropertyMetadata(null));

        

        private string[] includedImages;

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void Set<T>(ref T value, T field, [CallerMemberName] string PropertyName = null)
        {
            if (!EqualityComparer<T>.Default.Equals(value, field))
            {
                field = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(PropertyName));
            }
        }
        private void RefreshUI()
        {
            currentSlide = 0;
            if (ItemSource == null || ItemSource.Count() == 0)
            {
                maximumSlides = 0;
                return;
            }

            maximumSlides = ItemSource.Where(x => x.IsIncluded)
                                      .Sum(x => x.PreviewImageRange.Item2 - x.PreviewImageRange.Item1);

            includedImages = ItemSource.Where(x => x.IsIncluded)
                                       .SelectMany(x => Enumerable.Range(x.PreviewImageRange.Item1, x.PreviewImageRange.Item2))
                                       .Select(x => System.IO.Path.Combine(tempPath, $"{x}.png"))
                                       .ToArray();
        }
    }
}
