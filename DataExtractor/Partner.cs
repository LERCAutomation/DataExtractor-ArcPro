using System.ComponentModel;

namespace DataExtractor
{
    /// <summary>
    /// Partner to extract.
    /// </summary>
    public class Partner : INotifyPropertyChanged
    {
        #region Fields

        public string PartnerName { get; set; }

        public string ShortName { get; set; }

        public string GISFormat { get; set; }

        public string ExportFormat { get; set; }

        public string SQLTable { get; set; }

        public string SQLFiles { get; set; }

        public string MapFiles { get; set; }

        public string Tags { get; set; }

        public string Notes { get; set; }

        private bool _isSelected;

        public bool IsSelected
        {
            get { return _isSelected; }
            set
            {
                _isSelected = value;

                OnPropertyChanged(nameof(IsSelected));
            }
        }

        #endregion Fields

        #region Creator

        public Partner()
        {
            // constructor takes no arguments.
        }

        public Partner(string partnerName)
        {
            PartnerName = partnerName;
        }

        #endregion Creator

        #region INotifyPropertyChanged Members

        /// <summary>
        /// Raised when a property on this object has a new value.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Raises this object's PropertyChanged event.
        /// </summary>
        /// <param name="propertyName">The property that has a new value.</param>
        internal virtual void OnPropertyChanged(string propertyName)
        {
            //VerifyPropertyName(propertyName);

            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                PropertyChangedEventArgs e = new(propertyName);
                handler(this, e);
            }
        }

        #endregion INotifyPropertyChanged Members
    }
}