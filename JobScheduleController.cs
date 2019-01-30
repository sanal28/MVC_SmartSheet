using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using SmartSheet_V._1._0.CommonLibrary;

namespace SmartSheet_V._1._0.Controllers
{
    public class JobScheduleController : Controller
    {
        string consString = ConfigurationManager.ConnectionStrings["Connection"].ConnectionString;
        // GET: JobSchedule
        [SessionAuthorize]
        public ActionResult JobScheduling()
        {
            return View();
        }


        public JsonResult GetTaskName()
        {

            try
            {
                string jResult = string.Empty;
                using (SqlConnection con = new SqlConnection(consString))
                {
                    using (SqlCommand cmd = new SqlCommand("VTwo.UspDropDownBind"))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Connection = con;
                        cmd.Parameters.AddWithValue("@ID_Module", 4);
                        con.Open();
                        cmd.ExecuteNonQuery();
                        SqlDataReader dataReader = cmd.ExecuteReader();
                        DataTable dataTable = new DataTable();
                        dataTable.Load(dataReader);
                        con.Close();
                        if (dataTable.Rows.Count > 0)
                        {
                            // return (dataTable);
                            jResult = JsonConvert.SerializeObject(dataTable, Newtonsoft.Json.Formatting.Indented);
                            if (jResult != string.Empty)
                                return Json(jResult, JsonRequestBehavior.AllowGet);
                        }
                        else
                        {

                        }

                    }
                }

            }
            catch (Exception ex)
            {

            }
            finally
            {
                Dispose();
            }
            return Json(new { flag = false }, JsonRequestBehavior.AllowGet);
        }


        [HttpPost]
       
        public JsonResult WeeklyScheduleTask(int Weeekdays, int taskName, 
            int Occurs, int monthlyRecurs,
           int occursOnce, string occursoncetime, string occurseverystartTime, string occurseveryEndtime, int occurseveyMinutes, int occurseveryText,
            string startDate, string endDate, int radioendoption)
        {
            string CornExpression = string.Empty;
            string weekday = "";
            string []_timeSpilt;
            string[] _StarttimeSpilt;
            string[] _EndtimeSpilt;
            int _occurseveyMinutes=0, _hours=0, _mintues, _seconds;
            string[] _startdateSplit, _EnddateSplit;
            DateTime startdt = Convert.ToDateTime(startDate);
            string newStartDateString = startdt.ToString("dd-MMM-yy");
            DateTime enddt = Convert.ToDateTime(endDate);
            string newEndDateString = enddt.ToString("dd-MMM-yy");
            try
            {

                if (((Weeekdays & 1) > 0))
                    weekday += "SUN";

                if (((Weeekdays & 2) > 0))

                    weekday += " MON";
                if (((Weeekdays & 4) > 0))
                    weekday += " TUE";

                if (((Weeekdays & 8) > 0))
                    weekday += " WED";

                if (((Weeekdays & 16) > 0))
                    weekday += " THU";

                if (((Weeekdays & 32) > 0))
                    weekday += " FRI";

                if (((Weeekdays & 64) > 0))
                    weekday += " SAT";
                weekday = (weekday.Replace(' ', ',')).TrimStart(',');
                CornExpression="0 0 0 ? * "+ weekday + " *";
               

                if (occursOnce==1)
                {
                    _timeSpilt = occursoncetime.Split(':');
                    CornExpression ="0 " + _timeSpilt [1]+ " "+ _timeSpilt[0] + " ? * " + weekday + " *";

                }
                else
                {
                    _StarttimeSpilt= occurseverystartTime.Split(':');
                    _EndtimeSpilt = occurseveryEndtime.Split(':');
                    CornExpression = "0 " + _StarttimeSpilt[1] +"-" + _EndtimeSpilt[1] + " " + _StarttimeSpilt[0] + "-" + _EndtimeSpilt[0]+ " ? * " + weekday + " *";
                    if (occurseveryText == 0)
                    {
                        //hours
                        _hours = occurseveyMinutes;

                    }
                    else if (occurseveryText == 1)
                    {
                        //Minutes
                        _mintues = occurseveyMinutes;

                    }
                    else
                    {
                        //seconds
                        _seconds = occurseveyMinutes;
                    }
                }
                // _occurseveyMinutes = occurseveyMinutes;
                _startdateSplit = newStartDateString.Split('-');
                if (radioendoption==1)
                {
                    _EnddateSplit = "00-00-0000".Split('-');
                }
                else
                {
                    _EnddateSplit = newEndDateString.Split('-');
                }
                

                return Json(new { id = 1 }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                
                return Json(new { id = -1 }, JsonRequestBehavior.AllowGet);
            }
            finally
            {

                Dispose();
            }
           // return Json(CommonLibrary.Constants.JsonError, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        
        public JsonResult MonthlyScheduleTask(int packageName,int Occurs, int radioMonthlyvalueoption,
      int optionDayevery, int optiondayMonths, int optiontheDay, int optionTheweek, int optiontheMonths, int occursOnce,
      string occursoncetime, string occurseverystartTime, string occurseveryEndtime, int occurseveyMinutes, int occurseveryText,
      string startDate, string endDate, int radioendoption
  )
        {
            try
            {
                //AssignScheduleTaskMonthly(Name, Occurs, radioMonthlyvalueoption, optionDayevery, optiondayMonths,
                //    optiontheDay, optionTheweek, optiontheMonths, occursOnce, occursoncetime, occurseveryText, occurseveyMinutes,
                //    occurseverystartTime, occurseveryEndtime, startDate, endDate, radioendoption, packageName, packagePath);
                //int pkgname = packageName;
                return Json(new { id = 1 }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
               
                return Json(new { id = -1 }, JsonRequestBehavior.AllowGet);
            }
            finally
            {

                Dispose();
            }

        }
    }
}