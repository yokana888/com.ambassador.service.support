using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using com.ambassador.support.lib.Interfaces;
using com.ambassador.support.lib.Services;
using com.ambassador.support.webapi.Helpers;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace com.ambassador.support.webapi.Controllers.v1
{
	[Produces("application/json")]
	[ApiVersion("1.0")]
	[Route("v{version:apiVersion}/customs-reports")]
	[Authorize]
	public class CustomsReportController : Controller
	{
		private static readonly string ApiVersion = "1.0";
		private ScrapService scrapService { get; }
        //private FactBeacukaiService factBeacukaiService { get; }
        private FactItemMutationService factItemMutationService { get; }
        private WIPService wipService { get; }
		private FinishedGoodService finishedGoodService { get; }
		private MachineMutationService machineMutationService { get; }
        private HOrderService hOrderService { get; }
        private ExpenditureGoodsService expenditureGoodsService { get; }
        private TraceableInService traceableInService { get; }
        private TraceableOutService traceableOutService { get; }
        private readonly IExpenditureRawMaterialService expenditureRawMaterialService;
        private readonly IReceiptRawMaterialService receiptRawMaterialService;
        private readonly IFinishingOutOfGoodService finishingOutOfGoodService;
        private readonly IWasteScrapService wasteScrapService;
        private readonly IWIPInSubconService wIPInSubconService;
        private readonly IFactBeacukaiService factBeacukaiService;
        private IdentityService IdentityService;


        public CustomsReportController(IExpenditureRawMaterialService expenditureRawMaterialService, IReceiptRawMaterialService receiptRawMaterialService, IFinishingOutOfGoodService finishingOutOfGoodService, IWasteScrapService wasteScrapService, IWIPInSubconService wIPInSubconService, IFactBeacukaiService factBeacukaiService, IdentityService identityService )
        {
			this.scrapService = scrapService;
            this.factBeacukaiService = factBeacukaiService;
            this.factItemMutationService = factItemMutationService;
			this.wipService = wipService;
			this.finishedGoodService = finishedGoodService;
			this.machineMutationService = machineMutationService;
            this.hOrderService = hOrderService;
            this.expenditureGoodsService = expenditureGoodsService;
            this.traceableInService = traceableInService;
            this.traceableOutService = traceableOutService;
            this.expenditureRawMaterialService = expenditureRawMaterialService;
            this.receiptRawMaterialService = receiptRawMaterialService;
            this.finishingOutOfGoodService = finishingOutOfGoodService;
            this.wasteScrapService = wasteScrapService;
            this.wIPInSubconService = wIPInSubconService;
            IdentityService = identityService;
        }

        [HttpGet("expenditure-raw-material")]
        public async Task<IActionResult> GetExpenditureRawMaterial(DateTimeOffset? dateFrom, DateTimeOffset? dateTo, int page, int size, string Order = "{}")
        {
            int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
            string accept = Request.Headers["Accept"];
            IdentityService.Token = Request.Headers["Authorization"].FirstOrDefault().Replace("Bearer ", "");
            try
            {

                var data =await expenditureRawMaterialService.GetReport(dateFrom, dateTo, page, size, Order, offset);

                return Ok(new
                {
                    apiVersion = ApiVersion,
                    data = data.Item1,
                    info = new { total = data.Item2 }
                });
            }
            catch (Exception e)
            {
                Dictionary<string, object> Result =
                    new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
                    .Fail();
                return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
            }
        }

        [HttpGet("expenditure-raw-material/download")]
        public async Task<IActionResult> GetXlsIN( DateTimeOffset? dateFrom, DateTimeOffset? dateTo)
        {
            int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
            string accept = Request.Headers["Accept"];
            IdentityService.Token = Request.Headers["Authorization"].FirstOrDefault().Replace("Bearer ", "");
            try
            {
                byte[] xlsInBytes;
                //DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : Convert.ToDateTime(dateFrom);
                //DateTime DateTo = dateTo == null ? DateTime.Now : Convert.ToDateTime(dateTo);

                var xls = await expenditureRawMaterialService.GenerateExcel(dateFrom, dateTo, offset);

                string filename = String.Format("Laporan Pemakaian Bahan Baku - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

                xlsInBytes = xls.ToArray();
                var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
                return file;

            }
            catch (Exception e)
            {
                Dictionary<string, object> Result =
                    new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
                    .Fail();
                return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
            }
        }

        [HttpGet("receipt-raw-material")]
        public async Task<IActionResult> GetReceiptRawMaterial(DateTime? dateFrom, DateTime? dateTo, int page, int size, string Order = "{}")
        {
            

            try
            {
                int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
                string accept = Request.Headers["Accept"];
                IdentityService.Token = Request.Headers["Authorization"].FirstOrDefault().Replace("Bearer ", "");
                var data = await receiptRawMaterialService.GetReport(dateFrom, dateTo, page, size, Order);

                return Ok(new
                {
                    apiVersion = ApiVersion,
                    data = data.Item1,
                    info = new { total = data.Item2 }
                });
            }
            catch (Exception e)
            {
                Dictionary<string, object> Result =
                    new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
                    .Fail();
                return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
            }
        }

        [HttpGet("receipt-raw-material/download")]
        public async Task<IActionResult> GetExcelRawMaterial(DateTime? dateFrom, DateTime? dateTo)
        {
            try
            {
                byte[] xlsInBytes;

                int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
                string accept = Request.Headers["Accept"];
                IdentityService.Token = Request.Headers["Authorization"].FirstOrDefault().Replace("Bearer ", "");

                var xls = await receiptRawMaterialService.GenerateExcel(dateFrom, dateTo);

                string filename = String.Format("Laporan Pemasukan Bahan Baku - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

                xlsInBytes = xls.ToArray();
                var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
                return file;

            }
            catch (Exception e)
            {
                Dictionary<string, object> Result =
                    new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
                    .Fail();
                return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
            }
        }

        [HttpGet("finishing-out-of-good")]
        public IActionResult GetFinisingOutOfGood(DateTime? dateFrom, DateTime? dateTo, int page, int size, string Order = "{}")
        {
            int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
            string accept = Request.Headers["Accept"];

            try
            {
                var data = finishingOutOfGoodService.GetReport(dateFrom, dateTo, page, size, Order);

                return Ok(new
                {
                    apiVersion = ApiVersion,
                    data = data.Item1,
                    info = new { total = data.Item2 }
                });
            }
            catch (Exception e)
            {
                Dictionary<string, object> Result =
                    new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
                    .Fail();
                return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
            }
        }

        [HttpGet("finishing-out-of-good/download")]
        public IActionResult GetExcelFinisingOutOfGood(DateTime? dateFrom, DateTime? dateTo)
        {
            int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
            string accept = Request.Headers["Accept"];
            try
            {
                byte[] xlsInBytes;

                var xls = finishingOutOfGoodService.GenerateExcel(dateFrom, dateTo);

                string filename = String.Format("Laporan Pemasukan Barang Jadi - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

                xlsInBytes = xls.ToArray();
                var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
                return file;

            }
            catch (Exception e)
            {
                Dictionary<string, object> Result =
                    new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
                    .Fail();
                return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
            }
        }

        [HttpGet("waste-scrap")]
        public IActionResult GetWasteScrap(DateTime? dateFrom, DateTime? dateTo, int page, int size, string Order = "{}")
        {
            int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
            string accept = Request.Headers["Accept"];

            try
            {
                var data = wasteScrapService.GetReport(dateFrom, dateTo, page, size, Order);

                return Ok(new
                {
                    apiVersion = ApiVersion,
                    data = data.Item1,
                    info = new { total = data.Item2 }
                });
            }
            catch (Exception e)
            {
                Dictionary<string, object> Result =
                    new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
                    .Fail();
                return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
            }
        }

        [HttpGet("waste-scrap/download")]
        public IActionResult GetExcelWasteScrap(DateTime? dateFrom, DateTime? dateTo)
        {
            int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
            string accept = Request.Headers["Accept"];
            try
            {
                byte[] xlsInBytes;

                var xls = wasteScrapService.GenerateExcel(dateFrom, dateTo);

                string filename = String.Format("Laporan Penyelesaian Waste Scrap - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

                xlsInBytes = xls.ToArray();
                var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
                return file;

            }
            catch (Exception e)
            {
                Dictionary<string, object> Result =
                    new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
                    .Fail();
                return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
            }
        }

        [HttpGet("wip-in-subcon")]
        public IActionResult GetWipInSubcon(DateTimeOffset? dateFrom, DateTimeOffset? dateTo, int page, int size, string Order = "{}")
        {
            int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
            string accept = Request.Headers["Accept"];

            try
            {

                var data = wIPInSubconService.GetReport(dateFrom, dateTo, page, size, Order, offset);

                return Ok(new
                {
                    apiVersion = ApiVersion,
                    data = data.Item1,
                    info = new { total = data.Item2 }
                });
            }
            catch (Exception e)
            {
                Dictionary<string, object> Result =
                    new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
                    .Fail();
                return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
            }
        }

        [HttpGet("wip-in-subcon/download")]
        public IActionResult GetXlsWipInSubcon(DateTimeOffset? dateFrom, DateTimeOffset? dateTo)
        {
            int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
            string accept = Request.Headers["Accept"];
            try
            {
                byte[] xlsInBytes;
                //DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : Convert.ToDateTime(dateFrom);
                //DateTime DateTo = dateTo == null ? DateTime.Now : Convert.ToDateTime(dateTo);

                var xls = wIPInSubconService.GenerateExcel(dateFrom, dateTo, offset);

                string filename = String.Format("Laporan Pemakaian Barang Dalam Proses Subkontrak - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

                xlsInBytes = xls.ToArray();
                var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
                return file;

            }
            catch (Exception e)
            {
                Dictionary<string, object> Result =
                    new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
                    .Fail();
                return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
            }
        }

        [HttpGet("getPEB")]
        public IActionResult GetPEB([FromBody] string invoice)
        {
            int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
            string accept = Request.Headers["Accept"];

            try
            {

                var data = factBeacukaiService.GetBEACUKAI_ADDEDs(invoice);

                return Ok(new
                {
                    apiVersion = ApiVersion,
                    data = data,
                    info = new { total = data.Count() }
                });
            }
            catch (Exception e)
            {
                Dictionary<string, object> Result =
                    new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
                    .Fail();
                return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
            }
        }

        //[HttpGet("in")]
        //public IActionResult GetIN(string type, DateTime? dateFrom, DateTime? dateTo, int page, int size, string Order = "{}")
        //{
        //    int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //    string accept = Request.Headers["Accept"];

        //    try
        //    {

        //        var data = factBeacukaiService.GetReportIN(type, dateFrom, dateTo, page, size, Order, offset);

        //        return Ok(new
        //        {
        //            apiVersion = ApiVersion,
        //            data = data.Item1,
        //            info = new { total = data.Item2 }
        //        });
        //    }
        //    catch (Exception e)
        //    {
        //        Dictionary<string, object> Result =
        //            new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //            .Fail();
        //        return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //    }
        //}

        //      [HttpGet("in/download")]
        //      public IActionResult GetXlsIN(string type, DateTime? dateFrom, DateTime? dateTo)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];
        //          try
        //          {
        //              byte[] xlsInBytes;
        //              DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : Convert.ToDateTime(dateFrom);
        //              DateTime DateTo = dateTo == null ? DateTime.Now : Convert.ToDateTime(dateTo);

        //              var xls = factBeacukaiService.GenerateExcelIN(type, dateFrom, dateTo, offset);

        //              string filename = String.Format("Laporan Pemasukan Barang per Dokumen Pabean - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

        //              xlsInBytes = xls.ToArray();
        //              var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        //              return file;

        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("out")]
        //      public IActionResult GetOUT(string type, DateTime? dateFrom, DateTime? dateTo, int page, int size, string Order = "{}")
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];

        //          try
        //          {

        //              var data = factBeacukaiService.GetReportOUT(type, dateFrom, dateTo, page, size, Order, offset);

        //              return Ok(new
        //              {
        //                  apiVersion = ApiVersion,
        //                  data = data.Item1,
        //                  info = new { total = data.Item2 }
        //              });
        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("out/download")]
        //      public IActionResult GetXlsOUT(string type, DateTime? dateFrom, DateTime? dateTo)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];
        //          try
        //          {
        //              byte[] xlsInBytes;
        //              DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : Convert.ToDateTime(dateFrom);
        //              DateTime DateTo = dateTo == null ? DateTime.Now : Convert.ToDateTime(dateTo);

        //              var xls = factBeacukaiService.GenerateExcelOUT(type, dateFrom, dateTo, offset);

        //              string filename = String.Format("Laporan Pengeluaran Barang per Dokumen Pabean - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

        //              xlsInBytes = xls.ToArray();
        //              var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        //              return file;

        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("scrap")]
        //public IActionResult Get(  DateTime? dateFrom, DateTime? dateTo)
        //{
        //	int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //	string accept = Request.Headers["Accept"];

        //	try
        //	{
        //		var data = scrapService.GetScrapReport(dateFrom,dateTo, offset);
        //		return Ok(new
        //		{
        //			apiVersion = ApiVersion,
        //			data = data
        //		});
        //	}
        //	catch (Exception e)
        //	{
        //		Dictionary<string, object> Result =
        //			new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //			.Fail();
        //		return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //	}
        //}

        //      [HttpGet("scrap/download")]
        //      public IActionResult GetXlsScrap(DateTime? dateFrom, DateTime? dateTo)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];
        //          try
        //          {
        //              byte[] xlsInBytes;
        //              DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : Convert.ToDateTime(dateFrom);
        //              DateTime DateTo = dateTo == null ? DateTime.Now : Convert.ToDateTime(dateTo);

        //              var xls = scrapService.GenerateExcel(dateFrom, dateTo, offset);

        //              string filename = String.Format("Laporan Pertanggungjawaban Barang Reject dan Scrap - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

        //              xlsInBytes = xls.ToArray();
        //              var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        //              return file;

        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("finished-good")]
        //public IActionResult GetFinishedGood(DateTime? dateFrom, DateTime? dateTo)
        //{
        //	int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //	string accept = Request.Headers["Accept"];

        //	try
        //	{
        //		var data = finishedGoodService.GetFinishedGoodReport(dateFrom, dateTo, offset);
        //		return Ok(new
        //		{
        //			apiVersion = ApiVersion,
        //			data = data
        //		});
        //	}
        //	catch (Exception e)
        //	{
        //		Dictionary<string, object> Result =
        //			new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //			.Fail();
        //		return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //	}
        //}

        //      [HttpGet("finished-good/download")]
        //      public IActionResult GetXlsFinishedGood(DateTime? dateFrom, DateTime? dateTo)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];
        //          try
        //          {
        //              byte[] xlsInBytes;
        //              DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : Convert.ToDateTime(dateFrom);
        //              DateTime DateTo = dateTo == null ? DateTime.Now : Convert.ToDateTime(dateTo);

        //              var xls = finishedGoodService.GenerateExcel(dateFrom, dateTo, offset);

        //              string filename = String.Format("Laporan Pertanggungjawaban Mutasi Barang Jadi - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

        //              xlsInBytes = xls.ToArray();
        //              var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        //              return file;

        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("machine-mutation")]
        //public IActionResult GetMachineReport(DateTime? dateFrom, DateTime? dateTo, int page, int size, string Order = "{}")
        //{
        //	int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //	string accept = Request.Headers["Accept"];

        //	try
        //	{
        //		var data = machineMutationService.GetMachineMutationReportData(dateFrom, dateTo, page, size, Order, offset);
        //		return Ok(new
        //		{
        //                  apiVersion = ApiVersion,
        //                  data = data.Item1,
        //                  info = new { total = data.Item2 }
        //              });
        //	}
        //	catch (Exception e)
        //	{
        //		Dictionary<string, object> Result =
        //			new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //			.Fail();
        //		return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //	}
        //      }

        //      [HttpGet("machine-mutation/download")]
        //      public IActionResult GetXlsMachine(DateTime? dateFrom, DateTime? dateTo)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];
        //          try
        //          {
        //              byte[] xlsInBytes;
        //              DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : Convert.ToDateTime(dateFrom);
        //              DateTime DateTo = dateTo == null ? DateTime.Now : Convert.ToDateTime(dateTo);

        //              var xls = machineMutationService.GenerateExcel(dateFrom, dateTo, offset);

        //              string filename = String.Format("Laporan Pertanggungjawaban Mutasi Mesin dan Peralatan - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

        //              xlsInBytes = xls.ToArray();
        //              var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        //              return file;

        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("wip")]
        //      public IActionResult GetWIP(DateTime? date)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];

        //          try
        //          {
        //              var data = wipService.GetWIPReport(date, offset);
        //              return Ok(new
        //              {
        //                  apiVersion = ApiVersion,
        //                  data = data
        //              });
        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("wip/download")]
        //      public IActionResult GetXlsWIP(DateTime? date)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];
        //          try
        //          {
        //              byte[] xlsInBytes;

        //              var xls = wipService.GenerateExcel(date, offset);

        //              string filename = String.Format("Laporan Posisi WIP - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

        //              xlsInBytes = xls.ToArray();
        //              var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        //              return file;

        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("bbUnits")]
        //      public IActionResult GetBBUnit(int unit, DateTime? dateFrom, DateTime? dateTo, int page, int size, string Order = "{}")
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];

        //          try
        //          {
        //              var data = factItemMutationService.GetReportBBUnit(unit, dateFrom, dateTo, page, size, Order, offset);
        //              return Ok(new
        //              {
        //                  apiVersion = ApiVersion,
        //                  data = data.Item1,
        //                  info = new { total = data.Item2 }
        //              });
        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("bbUnits/download")]
        //      public IActionResult GetXlsBBUnit(int unit, DateTime? dateFrom, DateTime? dateTo)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];
        //          try
        //          {
        //              byte[] xlsInBytes;
        //              DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : Convert.ToDateTime(dateFrom);
        //              DateTime DateTo = dateTo == null ? DateTime.Now : Convert.ToDateTime(dateTo);

        //              var xls = factItemMutationService.GenerateExcelBBUnit(unit, dateFrom, dateTo, offset);

        //              string filename = String.Format("Laporan Pertanggungjawaban Mutasi Bahan Baku Unit - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

        //              xlsInBytes = xls.ToArray();
        //              var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        //              return file;

        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("bpUnits")]
        //      public IActionResult GetBPUnit(int unit, DateTime? dateFrom, DateTime? dateTo, int page, int size, string Order = "{}")
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];

        //          try
        //          {
        //              var data = factItemMutationService.GetReportBPUnit(unit, dateFrom, dateTo, page, size, Order, offset);
        //              return Ok(new
        //              {
        //                  apiVersion = ApiVersion,
        //                  data = data.Item1,
        //                  info = new { total = data.Item2 }
        //              });
        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("bpUnits/download")]
        //      public IActionResult GetXlsBPUnit(int unit, DateTime? dateFrom, DateTime? dateTo)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];
        //          try
        //          {
        //              byte[] xlsInBytes;
        //              DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : Convert.ToDateTime(dateFrom);
        //              DateTime DateTo = dateTo == null ? DateTime.Now : Convert.ToDateTime(dateTo);

        //              var xls = factItemMutationService.GenerateExcelBPUnit(unit, dateFrom, dateTo, offset);

        //              string filename = String.Format("Laporan Pertanggungjawaban Mutasi Bahan Penolong Unit - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

        //              xlsInBytes = xls.ToArray();
        //              var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        //              return file;

        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("bpCentrals")]
        //      public IActionResult GetBPCentral(DateTime? dateFrom, DateTime? dateTo, int page, int size, string Order = "{}")
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];

        //          try
        //          {
        //              var data = factItemMutationService.GetReportBPCentral(dateFrom, dateTo, page, size, Order, offset);
        //              return Ok(new
        //              {
        //                  apiVersion = ApiVersion,
        //                  data = data.Item1,
        //                  info = new { total = data.Item2 }
        //              });
        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("bpCentrals/download")]
        //      public IActionResult GetXlsBPCentral( DateTime? dateFrom, DateTime? dateTo)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];
        //          try
        //          {
        //              byte[] xlsInBytes;
        //              DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : Convert.ToDateTime(dateFrom);
        //              DateTime DateTo = dateTo == null ? DateTime.Now : Convert.ToDateTime(dateTo);

        //              var xls = factItemMutationService.GenerateExcelBPCentral(dateFrom, dateTo, offset);

        //              string filename = String.Format("Laporan Pertanggungjawaban Mutasi Bahan Penolong Pusat - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

        //              xlsInBytes = xls.ToArray();
        //              var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        //              return file;

        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("bbCentrals")]
        //      public IActionResult GetBBCentral(DateTime? dateFrom, DateTime? dateTo, int page, int size, string Order = "{}")
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];

        //          try
        //          {
        //              var data = factItemMutationService.GetReportBBCentral(dateFrom, dateTo, page, size, Order, offset);
        //              return Ok(new
        //              {
        //                  apiVersion = ApiVersion,
        //                  data = data.Item1,
        //                  info = new { total = data.Item2 }
        //              });
        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("bbCentrals/download")]
        //      public IActionResult GetXlsBBCentral(DateTime? dateFrom, DateTime? dateTo)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];
        //          try
        //          {
        //              byte[] xlsInBytes;
        //              DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : Convert.ToDateTime(dateFrom);
        //              DateTime DateTo = dateTo == null ? DateTime.Now : Convert.ToDateTime(dateTo);

        //              var xls = factItemMutationService.GenerateExcelBBCentral(dateFrom, dateTo, offset);

        //              string filename = String.Format("Laporan Pertanggungjawaban Mutasi Bahan Baku Pusat - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

        //              xlsInBytes = xls.ToArray();
        //              var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        //              return file;

        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //[HttpGet("beacukaitemp")]
        //public IActionResult GetBeacukai(string Keyword = "", string Filter = "{}")
        //{
        //	int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //	string accept = Request.Headers["Accept"];

        //	try
        //	{
        //		var data = factBeacukaiService.GetBEACUKAI_TEMPs(Keyword,Filter);
        //		return Ok(new
        //		{
        //			apiVersion = ApiVersion,
        //			data = data
        //		});
        //	}
        //	catch (Exception e)
        //	{
        //		Dictionary<string, object> Result =
        //			new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //			.Fail();
        //		return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //	}
        //}

        //      [HttpGet("horder")]
        //      public IActionResult GetHOrder(string Keyword = "", string Filter = "{}")
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];

        //          try
        //          {
        //              var data = hOrderService.GetHOrders(Keyword, Filter);
        //              return Ok(new
        //              {
        //                  apiVersion = ApiVersion,
        //                  data = data
        //              });
        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("exgood")]
        //      public IActionResult GetExGood(DateTime? dateFrom, DateTime? dateTo, int page, int size, string Order = "{}")
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];

        //          try
        //          {
        //              var data = expenditureGoodsService.GetReportExGood(dateFrom, dateTo, page, size, Order, offset);
        //              return Ok(new
        //              {
        //                  apiVersion = ApiVersion,
        //                  data = data.Item1,
        //                  info = new { total = data.Item2 }
        //              });
        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }
        //      [HttpGet("exgood/download")]
        //      public IActionResult GetXlsIN(DateTime? dateFrom, DateTime? dateTo)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];
        //          try
        //          {
        //              byte[] xlsInBytes;
        //              DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : Convert.ToDateTime(dateFrom);
        //              DateTime DateTo = dateTo == null ? DateTime.Now : Convert.ToDateTime(dateTo);

        //              var xls = expenditureGoodsService.GenerateExcelExGood(dateFrom, dateTo, offset);

        //              string filename = String.Format("Laporan Pengeluaran Barang Jadi - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

        //              xlsInBytes = xls.ToArray();
        //              var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        //              return file;

        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }
        //      [HttpGet("traceable/in")]
        //      public IActionResult GettraceIn(string bcno, string type, string tipebc, DateTime? DateFrom, DateTime? DateTo)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];

        //          try
        //          {
        //              var data2 = traceableInService.getQueryTracable2(bcno, type, tipebc, DateFrom, DateTo);
        //              return Ok(new
        //              {
        //                  apiVersion = ApiVersion,
        //                  data = data2,

        //              });
        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }
        //      [HttpGet("traceable/master")]
        //      public IActionResult Gettracemaster(int page = 1, int size = 25, string order = "{}", string keyword = null, string filter = "{}")
        //      {
        //          try
        //          {
        //              //identityService.Username = User.Claims.Single(p => p.Type.Equals("username")).Value;

        //              var model = traceableInService.Read(page, size, order, keyword, filter);

        //              var info = new Dictionary<string, object>
        //                  {
        //                      { "count", model.Data.Count },
        //                      { "total", model.TotalData },
        //                      { "order", model.Order },
        //                      { "page", page },
        //                      { "size", size }
        //                  };

        //              //Dictionary<string, object> Result =
        //              //    new ResultFormatter(ApiVersion, General.OK_STATUS_CODE, General.OK_MESSAGE)
        //              //    .Ok(model.Data, info);
        //              return Ok(new
        //              {
        //                  apiVersion = ApiVersion,
        //                  data = model.Data,
        //                  info = new Dictionary<string, object>
        //                  {
        //                      { "count", model.Data.Count },
        //                      { "total", model.TotalData },
        //                      { "order", model.Order },
        //                      { "page", page },
        //                      { "size", size }
        //                  }
        //              });
        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }
        //      [HttpGet("traceable/in/download")]
        //      public IActionResult GetXlsINTraceable(string bcno, string type, string tipebc, DateTime? DateFrom, DateTime? DateTo)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];
        //          try
        //          {
        //              byte[] xlsInBytes;
        //              //DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : Convert.ToDateTime(dateFrom);
        //              //DateTime DateTo = dateTo == null ? DateTime.Now : Convert.ToDateTime(dateTo);

        //              var xls = traceableInService.GetTraceableInExcel(bcno, type, tipebc, DateFrom, DateTo);

        //              string filename = String.Format("Laporan Traceable Masuk - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

        //              xlsInBytes = xls.ToArray();
        //              var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        //              return file;

        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("traceable/out")]
        //      public IActionResult GettraceOut(string bcno)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];

        //          try
        //          {
        //              var data2 = traceableOutService.getQueryTraceableOut(bcno);
        //              return Ok(new
        //              {
        //                  apiVersion = ApiVersion,
        //                  data = data2



        //              });
        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("traceableout/download")]
        //      public IActionResult GetXlsOutTraceable(string bcno)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];
        //          try
        //          {
        //              byte[] xlsInBytes;
        //              //DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : Convert.ToDateTime(dateFrom);
        //              //DateTime DateTo = dateTo == null ? DateTime.Now : Convert.ToDateTime(dateTo);

        //              var xls = traceableOutService.GetTraceableOutExcel(bcno);

        //              string filename = String.Format("Laporan Traceable Keluar - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

        //              xlsInBytes = xls.ToArray();
        //              var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        //              return file;

        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }

        //      [HttpGet("traceable/out/detail")]
        //      public IActionResult GettraceOutDetail(string ro)
        //      {
        //          int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //          string accept = Request.Headers["Accept"];

        //          try
        //          {
        //              var data2 = traceableOutService.getQueryTraceableOutDetail(ro);
        //              return Ok(new
        //              {
        //                  apiVersion = ApiVersion,
        //                  data = data2,

        //              });
        //          }
        //          catch (Exception e)
        //          {
        //              Dictionary<string, object> Result =
        //                  new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //                  .Fail();
        //              return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //          }
        //      }


        [HttpGet("getPEB/byBCNo")]
        public IActionResult GetPEBBCNo(string bcno)
        {
            int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
            string accept = Request.Headers["Accept"];

            try
            {

                var data = factBeacukaiService.GetBEACUKAI_ADDEDbyBCNo(bcno);

                return Ok(new
                {
                    apiVersion = ApiVersion,
                    data = data,
                    info = new { total = 1 }
                });
            }
            catch (Exception e)
            {
                Dictionary<string, object> Result =
                    new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
                    .Fail();
                return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
            }
        }

        [HttpGet("getPEB/byDate")]
        public IActionResult GetPEBDate(DateTime? dateFrom, DateTime? dateTo)
        {
            int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
            string accept = Request.Headers["Accept"];

            try
            {

                var data = factBeacukaiService.GetBEACUKAI_ADDEDbyDate(dateFrom, dateTo);

                return Ok(new
                {
                    apiVersion = ApiVersion,
                    data = data,
                    info = new { total = 1 }
                });
            }
            catch (Exception e)
            {
                Dictionary<string, object> Result =
                    new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
                    .Fail();
                return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
            }
        }

        //[HttpGet("traceable/out/download")]
        //public IActionResult GetXlsOutTraceableDetail(string bcno, string type)
        //{
        //    int offset = Convert.ToInt32(Request.Headers["x-timezone-offset"]);
        //    string accept = Request.Headers["Accept"];
        //    try
        //    {
        //        byte[] xlsInBytes;
        //        //DateTime DateFrom = dateFrom == null ? new DateTime(1970, 1, 1) : Convert.ToDateTime(dateFrom);
        //        //DateTime DateTo = dateTo == null ? DateTime.Now : Convert.ToDateTime(dateTo);

        //       // var xls = traceableOutService.GetTraceableOutDetailExcel(bcno);

        //        string filename = String.Format("Laporan Traceable Masuk - {0}.xlsx", DateTime.UtcNow.ToString("ddMMyyyy"));

        //        xlsInBytes = xls.ToArray();
        //        var file = File(xlsInBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        //        return file;

        //    }
        //    catch (Exception e)
        //    {
        //        Dictionary<string, object> Result =
        //            new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //            .Fail();
        //        return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //    }
        //}

        //[HttpGet("traceableout/master")]
        //public IActionResult Gettraceoutmaster(int page = 1, int size = 25, string order = "{}", string keyword = null, string filter = "{}")
        //{
        //    try
        //    {
        //        //identityService.Username = User.Claims.Single(p => p.Type.Equals("username")).Value;

        //        var model = traceableOutService.Read(page, size, order, keyword, filter);

        //        var info = new Dictionary<string, object>
        //            {
        //                { "count", model.Data.Count },
        //                { "total", model.TotalData },
        //                { "order", model.Order },
        //                { "page", page },
        //                { "size", size }
        //            };

        //        //Dictionary<string, object> Result =
        //        //    new ResultFormatter(ApiVersion, General.OK_STATUS_CODE, General.OK_MESSAGE)
        //        //    .Ok(model.Data, info);
        //        return Ok(new
        //        {
        //            apiVersion = ApiVersion,
        //            data = model.Data,
        //            info = new Dictionary<string, object>
        //            {
        //                { "count", model.Data.Count },
        //                { "total", model.TotalData },
        //                { "order", model.Order },
        //                { "page", page },
        //                { "size", size }
        //            }
        //        });
        //    }
        //    catch (Exception e)
        //    {
        //        Dictionary<string, object> Result =
        //            new ResultFormatter(ApiVersion, General.INTERNAL_ERROR_STATUS_CODE, e.Message)
        //            .Fail();
        //        return StatusCode(General.INTERNAL_ERROR_STATUS_CODE, Result);
        //    }
        //}
    }
}