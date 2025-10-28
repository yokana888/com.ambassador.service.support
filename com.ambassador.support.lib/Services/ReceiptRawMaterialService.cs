using com.ambassador.support.lib.Helpers;
using com.ambassador.support.lib.Interfaces;
using com.ambassador.support.lib.ViewModel;
using Com.Moonlay.NetCore.Lib;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace com.ambassador.support.lib.Services
{
    public class ReceiptRawMaterialService : IReceiptRawMaterialService
    {
        IPurchasingDBContext context;
        public readonly IServiceProvider serviceProvider;
        public ReceiptRawMaterialService(IPurchasingDBContext _context, IServiceProvider serviceProvider)
        {
            this.context = _context;
            this.serviceProvider = serviceProvider;
        }
        public async Task<IQueryable<ReceiptRawMaterialViewModel>> getQuery(DateTime? dateFrom, DateTime? dateTo)
        {
            var d1 = dateFrom.Value.ToString("yyyy-MM-dd");
            var d2 = dateTo.Value.ToString("yyyy-MM-dd");
            var customCategory = "Fasilitas";


            List<ReceiptRawMaterialViewModel> reportData = new List<ReceiptRawMaterialViewModel>();

            try
            {
                string connectionString = APIEndpoint.PurchasingConnectionString;
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(
                        "declare @StartDate datetime = '" + d1 + "' declare @EndDate datetime = '" + d2 + "' " +
                        "select distinct e.CustomsType,e.BeacukaiNo,convert(date,dateadd(hour,7,e.BeacukaiDate)) as BCDate,f.URNNo,convert(date,dateadd(hour,7,f.ReceiptDate)) as URNDate,g.ProductCode,g.ProductName," +
                        "sum(g.SmallQuantity) as SmallQuantity,g.SmallUomUnit,a.DOCurrencyCode,sum(cast((g.PricePerDealUnit * g.ReceiptQuantity) as decimal(18,2))) as Amount,a.SupplierName,a.Country, c.ProductSeries, c.DeletedAgent " +
                        "from GarmentDeliveryOrders a join GarmentDeliveryOrderItems b on a.id=b.GarmentDOId join GarmentDeliveryOrderDetails c on b.id=c.GarmentDOItemId " +
                        "join GarmentBeacukaiItems d on d.GarmentDOId=a.id join GarmentBeacukais e on e.id=d.BeacukaiId " +
                        "join GarmentUnitReceiptNoteItems g on c.id=g.DODetailId join GarmentUnitReceiptNotes f on g.URNId=f.Id " +
                        "where e.BeacukaiDate between @StartDate and @EndDate and a.CustomsCategory = '"+ customCategory + "' and f.URNType='PEMBELIAN' " +
                        "and a.IsDeleted=0 and b.IsDeleted=0 and c.IsDeleted=0 and d.IsDeleted=0 and e.IsDeleted=0 and f.IsDeleted=0 and g.IsDeleted=0 " +
                        "group by e.CustomsType,e.BeacukaiNo,e.BeacukaiDate,f.URNNo,f.ReceiptDate,g.ProductCode,g.ProductName,g.SmallUomUnit,a.DOCurrencyCode,a.SupplierName,a.Country,c.ProductSeries,c.DeletedAgent " +
                        "order by BCDate asc", conn))

                    {
                        SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                        DataSet dSet = new DataSet();
                        dataAdapter.Fill(dSet);
                        foreach (DataRow data in dSet.Tables[0].Rows)
                        {
                            ReceiptRawMaterialViewModel view = new ReceiptRawMaterialViewModel
                            {
                                CustomsType = data["CustomsType"].ToString(),
                                BeacukaiNo = data["BeacukaiNo"].ToString(),
                                BeacukaiDate = data["BCDate"].ToString(),
                                SerialNo = data["ProductSeries"].ToString(),
                                URNNo = data["URNNo"].ToString(),
                                URNDate = data["URNDate"].ToString(),
                                ProductCode = data["ProductCode"].ToString(),
                                ProductName = data["ProductName"].ToString(),
                                SmallUomUnit = data["SmallUomUnit"].ToString(),
                                SmallQuantity = (decimal)data["SmallQuantity"],
                                DOCurrencyCode = data["DOCurrencyCode"].ToString(),
                                Amount = (decimal)data["Amount"],
                                StorageName = "GUDANG AG2",
                                SupplierName = "-",
                                Country = data["Country"].ToString(),
                                DeletedAgent = data["DeletedAgent"].ToString()
                            };
                            reportData.Add(view);
                        }
                    }
                    conn.Close();
                }
            }
            catch (SqlException ex)
            { 
            }

            var Codes = await GetProductCode(string.Join(",", reportData.Select(x => x.ProductCode).Distinct().ToList()));

            string[] exceptionBCNo = { "629905", "627663" , "038117", "046380", "621904", "758615", "643895" };
            foreach(var a in reportData)
            {
                var trimProduct = a.ProductCode.Trim();
                var remark = Codes.FirstOrDefault(x => x.Code.Trim() == trimProduct);
                var Composition = remark == null ? "-" : remark.Composition;
                //var Width = remark == null ? "-" : remark.Width;
                //var Const = remark == null ? "-" : remark.Const;
                //var Yarn = remark == null ? "-" : remark.Yarn;
                //var Name = remark == null ? "-" : remark.Name;

                if (!exceptionBCNo.Contains(a.BeacukaiNo))
                {
                    a.ProductName = remark != null ? string.Concat(/*a.ProductName, " - ",*/ Composition) : a.ProductName;
                }
                else
                {
                    a.ProductName = string.Concat(/*a.ProductName, " - ",*/ Composition + " - "+ a.DeletedAgent);
                }
            }

            //Order by SerialNo
            var groupedData = reportData
                .GroupBy(x => new { x.BeacukaiNo, x.BeacukaiDate, x.CustomsType })
                .Select(g => new
                {
                    g.Key.BeacukaiNo,
                    g.Key.BeacukaiDate,
                    g.Key.CustomsType,
                    // Custom sorting untuk SerialNo string / angka
                    Items = g.OrderBy(i =>
                    {
                        // Jika bisa parse ke int, urutkan sebagai angka
                        return int.TryParse(i.SerialNo, out int n) ? n : int.MaxValue;
                    })
            .ThenBy(i => i.SerialNo) // jika bukan angka, urutkan alfabet
            .ToList()
                })
                .ToList();

            // Flatten data sesuai urutan group dan SerialNo
            var flattenedData = groupedData
                .SelectMany(g => g.Items)
                .ToList();

            return flattenedData.AsQueryable();
        }
        public async Task<Tuple<List<ReceiptRawMaterialViewModel>, int>> GetReport(DateTime? dateFrom, DateTime? dateTo, int page, int size, string Order)
        {
            var Query = await getQuery(dateFrom, dateTo);

            Dictionary<string, string> OrderDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(Order);
            if (OrderDictionary.Count.Equals(0))
            {
                //Query = Query.OrderBy(b => b.BeacukaiDate);
            }
            else
            {
                string Key = OrderDictionary.Keys.First();
                string OrderType = OrderDictionary[Key];

                //Query = Query.OrderBy(string.Concat(Key, " ", OrderType));
            }

            Pageable<ReceiptRawMaterialViewModel> pageable = new Pageable<ReceiptRawMaterialViewModel>(Query, page - 1, size);
            List<ReceiptRawMaterialViewModel> Data = pageable.Data.ToList<ReceiptRawMaterialViewModel>();

            int TotalData = pageable.TotalCount;

            return Tuple.Create(Data, TotalData);
        }

        public async Task<MemoryStream> GenerateExcel(DateTime? dateFrom, DateTime? dateTo)
        {
            var Query = await getQuery(dateFrom, dateTo);
            //Query = Query.OrderBy(b => b.BeacukaiDate);
            DataTable result = new DataTable();
            result.Columns.Add(new DataColumn() { ColumnName = "No", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Tgl Rekam", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Jenis Dokumen", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "No Bea Cukai", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Tgl Bea Cukai", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Kode HS", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Nomor Seri Barang", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "No Bukti Penerimaan", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Tgl Bukti Penerimaan", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Kode Barang", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Nama Barang", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Satuan", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Jumlah Terima", DataType = typeof(double) });
            result.Columns.Add(new DataColumn() { ColumnName = "Mata Uang", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Nilai Barang", DataType = typeof(double) });
            result.Columns.Add(new DataColumn() { ColumnName = "Gudang", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Penerima Sub Kontrak", DataType = typeof(String) });
            result.Columns.Add(new DataColumn() { ColumnName = "Negara Asal Barang", DataType = typeof(String) });

            if (Query.ToArray().Count() == 0)
            {
                result.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "", 0, "", 0, "", "", ""); // to allow column name to be generated properly for empty data as template
            }
            else
            {
                int i = 0;
                foreach (var item in Query)
                {
                    i++;
                    result.Rows.Add(i.ToString(), "-",item.CustomsType,item.BeacukaiNo,formattedDate(item.BeacukaiDate), "-",item.SerialNo,item.URNNo,formattedDate(item.URNDate),item.ProductCode,
                                    item.ProductName,item.SmallUomUnit,item.SmallQuantity,item.DOCurrencyCode,item.Amount,item.StorageName,item.SupplierName,item.Country);
                }
            }
            return Excel.CreateExcel(new List<KeyValuePair<DataTable, string>>() { new KeyValuePair<DataTable, string>(result, "Territory") }, true);

        }

        string formattedDate(string num)
        {
            DateTime date = DateTime.Parse(num);

            string datee = date.ToString("dd MMMM yyyy");
            

            return datee;
        }

        private async  Task<List<GarmentProductViewModel>> GetProductCode(string codes)
        {
            IHttpClientService httpClient = (IHttpClientService)this.serviceProvider.GetService(typeof(IHttpClientService));

            var garmentProductionUri = APIEndpoint.Core + $"master/garmentProducts/byCode?code=" + codes;

            var httpResponse = httpClient.GetAsync(garmentProductionUri).Result;
            if (httpResponse.IsSuccessStatusCode)
            {
                var content = httpResponse.Content.ReadAsStringAsync().Result;
                Dictionary<string, object> result = JsonConvert.DeserializeObject<Dictionary<string, object>>(content);

                List<GarmentProductViewModel> viewModel;
                if (result.GetValueOrDefault("data") == null)
                {
                    viewModel = new List<GarmentProductViewModel>();
                }
                else
                {
                    viewModel = JsonConvert.DeserializeObject<List<GarmentProductViewModel>>(result.GetValueOrDefault("data").ToString());

                }
                return viewModel;
            }
            else
            {
                List<GarmentProductViewModel> viewModel = new List<GarmentProductViewModel>();
                return viewModel;
            }
        }
    }
}
