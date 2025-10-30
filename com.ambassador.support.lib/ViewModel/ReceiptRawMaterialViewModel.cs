using System;
using System.Collections.Generic;
using System.Text;

namespace com.ambassador.support.lib.ViewModel
{
    public class ReceiptRawMaterialViewModel
    {
        public string CustomsType { get; set; }
        public string BeacukaiNo { get; set; }
        public string BeacukaiDate { get; set; }
        public string SerialNo { get; set; }
        public string URNNo { get; set; }
        public string URNDate { get; set; }
        public string ProductCode { get; set; }
        public string ProductName { get; set; }
        public decimal SmallQuantity { get; set; }
        public string SmallUomUnit { get; set; }
        public string DOCurrencyCode { get; set; }        
        public decimal Amount { get; set; }
        public string StorageName { get; set; }
        public string SupplierName { get; set; }
        public string Country { get; set; }        
        public string DeletedAgent { get; set; }      
        public string HsCode { get; set; }
        public string RecordDate { get; set; }
    }
}
