using DevExpress.Xpo;
using System;

namespace TrainSheet
{
    /// <summary>
    /// Натурный лист
    /// </summary>
    public class WagonSheet : XPCustomObject
    {

        public WagonSheet() : base()
        {
            // This constructor is used when an object is loaded from a persistent storage.
            // Do not place any code here.
        }

        public WagonSheet(Session session) : base(session)
        {
            // This constructor is used when an object is loaded from a persistent storage.
            // Do not place any code here.
        }

        public override void AfterConstruction()
        {
            base.AfterConstruction();
            // Place here your initialization code.
        }

        private int _Id;
        [Key]
        public int Id
        {
            get { return _Id; }
            set { SetPropertyValue(nameof(Id), ref _Id, value); }
        }

        private int _TrainNumber;
        /// <summary>
        /// Номер поезда
        /// </summary>
        public int TrainNumber
        {
            get { return _TrainNumber; }
            set { SetPropertyValue(nameof(TrainNumber), ref _TrainNumber, value); }
        }

        private string _TrainIndexCombined;
        public string TrainIndexCombined
        {
            get { return _TrainIndexCombined; }
            set { SetPropertyValue(nameof(TrainIndexCombined), ref _TrainIndexCombined, value); }
        }
       // private string _WagonNumber;
        /// <summary>
        /// Номер состава 
        /// </summary>
        public int WagonNumber
        {
            get { return int.Parse(TrainIndexCombined.Substring(TrainIndexCombined.IndexOf("_") + 1, 3)); }
           // set { SetPropertyValue(nameof(WagonNumber), ref _WagonNumber, TrainIndexCombined); }
        }

        private string _ToStationName;
        /// <summary>
        /// Наименование станции назначения
        /// </summary>
        public string ToStationName
        {
            get { return _ToStationName; }
            set { SetPropertyValue(nameof(ToStationName), ref _ToStationName, value); }
        }

        private string _FromStationName;
        /// <summary>
        /// Наименование станции отправления
        /// </summary>
        public string FromStationName
        {
            get { return _FromStationName; }
            set { SetPropertyValue(nameof(FromStationName), ref _FromStationName, value); }
        }

        private string _LastStationName;
        /// <summary>
        /// Наименование станции дислокации
        /// </summary>
        public string LastStationName
        {
            get { return _LastStationName; }
            set { SetPropertyValue(nameof(LastStationName), ref _LastStationName, value); }
        }


        [Association]
        public XPCollection<WagonList> WagonLists
        {
            get { return GetCollection<WagonList>(nameof(WagonLists)); }
        }
    }
}