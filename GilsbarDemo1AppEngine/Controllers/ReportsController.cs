using GilsbarDemo1.Models;
using GilsbarDemo1AppEngine.Models;
using GilsbarDemo1AppEngine.Web_Data;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;

namespace GilsbarDemo1AppEngine
{
    public class ReportsController : ApiController
    {

        // GET /api/reports/SampleReport
        [Route("api/reports/SampleReport")]
        [HttpGet]
        public IHttpActionResult SampleReport()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "SampleReport";
                report.Filename = "SampleReport";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.SampleReport = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // SampleReport()

        // GET /api/reports/BuildingsIndexPrinterFriendly
        [Route("api/reports/BuildingsIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult BuildingsIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "BuildingsIndexPrinterFriendly";
                report.Filename = "BuildingsIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.BuildingsIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // BuildingsIndexPrinterFriendly()

        // GET /api/reports/DepartmentsIndexPrinterFriendly
        [Route("api/reports/DepartmentsIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult DepartmentsIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "DepartmentsIndexPrinterFriendly";
                report.Filename = "DepartmentsIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.DepartmentsIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // DepartmentsIndexPrinterFriendly()

        // GET /api/reports/JobTasksIndexPrinterFriendly
        [Route("api/reports/JobTasksIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult JobTasksIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "JobTasksIndexPrinterFriendly";
                report.Filename = "JobTasksIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.JobTasksIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // JobTasksIndexPrinterFriendly()

        // GET /api/reports/ManagersIndexPrinterFriendly
        [Route("api/reports/ManagersIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult ManagersIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "ManagersIndexPrinterFriendly";
                report.Filename = "ManagersIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.ManagersIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // ManagersIndexPrinterFriendly()

        // GET /api/reports/PersonsIndexPrinterFriendly
        [Route("api/reports/PersonsIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult PersonsIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "PersonsIndexPrinterFriendly";
                report.Filename = "PersonsIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.PersonsIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // PersonsIndexPrinterFriendly()

    }
}

