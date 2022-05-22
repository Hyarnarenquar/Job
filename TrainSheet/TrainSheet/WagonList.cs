using DevExpress.Xpo;
using System;

namespace TrainSheet
{
    public class WagonList : XPCustomObject
    {
        public WagonList() : base()
        {
            // This constructor is used when an object is loaded from a persistent storage.
            // Do not place any code here.
        }

        public WagonList(Session session) : base(session)
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

        private int _PositionInTrain;
        /// <summary>
        /// Позиция вагона в составе
        /// </summary>
        public int PositionInTrain
        {
            get { return _PositionInTrain; }
            set { SetPropertyValue(nameof(PositionInTrain), ref _PositionInTrain, value); }
        }
       
        private int _CarNumber;
        /// <summary>
        /// Номер вагона
        /// </summary>
        public int CarNumber
        {
            get { return _CarNumber; }
            set { SetPropertyValue(nameof(CarNumber), ref _CarNumber, value); }
        }

        private string _InvoiceNum;
        /// <summary>
        /// Номер накладной
        /// </summary>
        public string InvoiceNum
        {
            get { return _InvoiceNum; }
            set { SetPropertyValue(nameof(InvoiceNum), ref _InvoiceNum, value); }
        }

        private DateTime _WhenLastOperation;
        /// <summary>
        /// Дата операции
        /// </summary>
        public DateTime WhenLastOperation
        {
            get { return _WhenLastOperation; }
            set { SetPropertyValue(nameof(WhenLastOperation), ref _WhenLastOperation, value); }
        }

        private string _LastOperationName;
        /// <summary>
        /// Последняя операция
        /// </summary>
         public string LastOperationName
        {
            get { return _LastOperationName; }
            set { SetPropertyValue(nameof(LastOperationName), ref _LastOperationName, value); }
        }

        private string _FreightEtsngName;
        /// <summary>
        /// Наименование груза
        /// </summary>
        public string FreightEtsngName
        {
            get { return _FreightEtsngName; }
            set { SetPropertyValue(nameof(FreightEtsngName), ref _FreightEtsngName,value); }
        }

        private double _FreightTotalWeightKg;
        /// <summary>
        /// Вес по документам
        /// </summary>
        public double FreightTotalWeightKg
        {
            get { return _FreightTotalWeightKg; }
            set { SetPropertyValue(nameof(FreightTotalWeightKg), ref _FreightTotalWeightKg, value); }
        }


        [Association]
        public WagonSheet WagonSheet
        {
            get { return _WagonSheet; }
            set { SetPropertyValue(nameof(WagonSheet), ref _WagonSheet, value); }
        }
        WagonSheet _WagonSheet;
    }
}