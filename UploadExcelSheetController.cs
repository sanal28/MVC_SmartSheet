using SmartSheet_V._1._0.CommonLibrary;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Newtonsoft.Json;

namespace SmartSheet_V._1._0.Controllers
{
    public class UploadExcelSheetController : Controller
    {
        string consString = ConfigurationManager.ConnectionStrings["Connection"].ConnectionString;
        // GET: UploadExcelSheet

 
        [HttpPost]
        
        public JsonResult UploadExcelFile()
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;

            try
            {
                var newImage = System.Web.HttpContext.Current.Request.Files["newImage"];
                var random = "";
                string path = null;
                if (newImage != null)
                {
                    random = DateTime.Now.ToString("ddMMyyhhmmss");
                    HttpPostedFile filebase = newImage;
                    var fileName = Path.GetFileName(filebase.FileName);
                    path = ("../Uploads/") + Path.GetFileNameWithoutExtension(fileName) + "^_^_^" + random + Path.GetExtension(fileName);
                    filebase.SaveAs(Server.MapPath("../Uploads/") + path);



                    filePath = Server.MapPath("../Uploads/") + path; //get the path of the file
                    fileExt = Path.GetExtension(filePath);//get the file extension


                    if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                    {
                        string myStr = string.Empty;
                        
                            DataTable dtExcel = new DataTable();
                            dtExcel = ReadExcel(filePath, fileExt);//read excel file
                            DataTable dtNewExcel = new DataTable();
                            dtNewExcel = dtExcel.Copy();
                            DataTable dta = new DataTable();
                            dta.Columns.Add("Name", typeof(System.String));
                            dta.Columns.Add("TeamName", typeof(System.String));
                            dta.Columns.Add("StartDate", typeof(System.String));
                            dta.Columns.Add("EndDate", typeof(System.String));
                            DataRow tblrow;
                             DataRow tblNewRow;// = dta.NewRow();
                        TimeSpan _offset = new TimeSpan(5, 30, 00);

                            string leapYear;

                            int days;

                            int EmployeeCount = Convert.ToInt32(ConfigurationManager.AppSettings["EmployeeCount"]);

                            Int16 ExcelColStart = Convert.ToInt16(ConfigurationManager.AppSettings["ExcelColumnStart"]);
                            //int ExcelRowStart
                            for (Int16 _rw = Convert.ToInt16(ConfigurationManager.AppSettings["ExcelRowStart"]); _rw <= EmployeeCount + 1; _rw++)
                            {
                                leapYear = dtExcel.Rows[0][2].ToString();
                                string[] currentYear = leapYear.Split('/');
                                days = DateTime.DaysInMonth(Convert.ToInt32(currentYear[2]), Convert.ToInt32(currentYear[1]));
                                for (Int16 col = ExcelColStart; col <= days + 1; col++) //for(Int16 _rw=2; _rw<=col; _rw++ )
                                {

                                    string empname = dtExcel.Rows[_rw][0].ToString();
                                    string teamName = dtExcel.Rows[_rw][1].ToString();

                                string dtval = dtExcel.Rows[0][col].ToString();


                                    //TimeSpan t = new TimeSpan();
                                    string[] _DoubleShift_time;

                                    string _time0 = dtExcel.Rows[_rw][col].ToString();
                                    if (_time0.Contains('\\'))
                                    {
                                        _DoubleShift_time = _time0.Split('\\');

                                    }
                                    else
                                    {
                                        _DoubleShift_time = new string[1];
                                        _DoubleShift_time[0] = _time0;
                                    }
                                    string _tempdtval = dtval;
                                    for (int i = 0; i < _DoubleShift_time.Length; i++)
                                    {
                                        DateTime EndDate;
                                        tblrow = dta.NewRow();
                                         tblNewRow = dta.NewRow();
                                    string[] dtary = _tempdtval.Split('/');
                                        EndDate = new DateTime(Convert.ToInt16(dtary[2]), Convert.ToInt16(dtary[1]), Convert.ToInt16(dtary[0]));
                                    DataRow dtnew;
                                    dtnew=Spilt_Function(_DoubleShift_time[i], dtval, empname, teamName, tblrow, _offset, EndDate);
                                    if(!dtnew.IsNull(0))
                                    {
                                        dta.Rows.Add(dtnew);
                                    }
                                   EndDate = EndDate.AddDays(1);
                                        _tempdtval = EndDate.ToString("dd/MM/yyyy");
                                    }
                                }

                            }
                          
                            if (dta.Rows.Count > 0)
                            {
                              
                                using (SqlConnection con = new SqlConnection(consString))
                                {
                                    using (SqlCommand cmd = new SqlCommand("VTwo.Insert_ShiftDetails"))
                                    {
                                        cmd.CommandType = CommandType.StoredProcedure;
                                        cmd.Connection = con;
                                        cmd.Parameters.AddWithValue("@ShiftDetails", dta);
                                        cmd.Parameters.AddWithValue("@FK_ModifiedBy", 1);
                                        cmd.Parameters.AddWithValue("@FK_EnteredBy", 1);
                                        
                                        con.Open();
                                        //cmd.ExecuteNonQuery();
                                        //SqlDataReader dataReader = cmd.ExecuteReader();
                                        DataTable dataTable = new DataTable();
                                        dataTable.Load(cmd.ExecuteReader());
                                       int Returnvalue = Convert.ToInt32(dataTable.Rows[0]["Returnvalue"]);
                                       con.Close();
                                    return Json(new { returnvalue = Returnvalue }, JsonRequestBehavior.AllowGet);
                               
                                }
                                }
                            }
                    }

                }
                
                return Json(new { flag = 1 }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
               
                return Json(new { flag = -1 }, JsonRequestBehavior.AllowGet);
            }
            finally
            {
                Dispose();
            }

        }

        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string ExcelName = System.IO.Path.GetFileNameWithoutExtension(fileName);
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)//compare the extension of the file
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';";
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from[Monthly$]", con);//here we read data from sheet1
                    oleAdpt.Fill(dtexcel);//fill excel data into dataTable
                }

                catch (Exception ex)
                {
                
                }
            }
            return dtexcel;
        }

        public DataRow Spilt_Function(string _time0, string dtval,
            string _EmpName,string _TeamName, DataRow dra, TimeSpan _offset, DateTimeOffset Enddt)
        {
            DateTimeOffset Startdt = new DateTimeOffset();
           
           // const string _constShiftTypeOFF = "OFF";
            
            string[] ary = dtval.Split('/');
            if (_time0.Length > 0)
            {
                switch (_time0.ToUpper())
                {
                    case "OFF":
                        _time0 = "23:59:59-23:59:59";
                        break;
                    case "IN-SHIFT":
                        _time0 = "00:00:01-23:59:59";
                        break;
                    case "COMP-OFF":
                    case "CO-OFF":
                        _time0 = "23:59:58-23:59:58";
                        break;
                    case "RH":
                        _time0 = "23:59:57-23:59:57";
                        break;
                    case "NH":
                        _time0 = "23:59:55-23:59:55";
                        break;
                    case "WH":
                        _time0 = "23:59:54-23:59:54";
                        break;
                    case "FH":
                        _time0 = "23:59:53-23:59:53";
                        break;
                    case "LEAVE":
                        _time0 = "23:59:56-23:59:56";
                        break;
                }
                string[] _time = _time0.Split('-');
                string[] _timetemp = _time[0].Split(':');
                Int32 _hr = 0, _min = 0, _sec = 0;
                if (_timetemp.Length == 2)
                {
                    _hr = Convert.ToInt32(_timetemp[0]);
                    _min = Convert.ToInt32(_timetemp[1]);
                }
                if (_timetemp.Length == 3)
                {
                    _hr = Convert.ToInt32(_timetemp[0]);
                    _min = Convert.ToInt32(_timetemp[1]);
                    _sec = Convert.ToInt32(_timetemp[2]);
                }
                Startdt = new DateTimeOffset(Convert.ToInt16(ary[2]), Convert.ToInt16(ary[1]), Convert.ToInt16(ary[0]),
                    _hr, _min, _sec, _offset);


                _timetemp = _time[1].Split(':');
                if (_timetemp.Length == 2)
                {
                    _hr = Convert.ToInt32(_timetemp[0]);
                    _min = Convert.ToInt32(_timetemp[1]);
                }
                if (_timetemp.Length == 3)
                {
                    _hr = Convert.ToInt32(_timetemp[0]);
                    _min = Convert.ToInt32(_timetemp[1]);
                    _sec = Convert.ToInt32(_timetemp[2]);
                }
                
                Enddt = new DateTimeOffset(Enddt.Year, Enddt.Month, Enddt.Day,
               _hr, _min, _sec, _offset);
                dra[0] = _EmpName;
                dra[1] = _TeamName;
                dra[2] = Startdt.ToString("yyyy-MM-dd HH:mm:ss zzz");
                dra[3] = Enddt.ToString("yyyy-MM-dd HH:mm:ss zzz");

            }
            return dra;
        }


        [HttpPost]
      
        public JsonResult GetShiftDetails(int MonthVal, int YearVal)
        {

            try
            {
                string jResult = string.Empty;
                using (SqlConnection con = new SqlConnection(consString))
                {
                    using (SqlCommand cmd = new SqlCommand("VTwo.UspRptMonthShift"))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Connection = con;
                        cmd.Parameters.AddWithValue("@ShiftMonth", MonthVal);
                        cmd.Parameters.AddWithValue("@ShiftYear ", YearVal);
                        cmd.Parameters.AddWithValue("@DataFormat", 1);
                        con.Open();
                        cmd.ExecuteNonQuery();
                        SqlDataReader dataReader = cmd.ExecuteReader();
                        DataTable dataTable = new DataTable();
                        dataTable.Load(dataReader);
                        con.Close();
                        if (dataTable.Rows.Count > 0)
                        {
                           // return (dataTable);
                            jResult= JsonConvert.SerializeObject(dataTable, Newtonsoft.Json.Formatting.Indented);
                            if (jResult != string.Empty)
                                return Json(new { result= jResult }, JsonRequestBehavior.AllowGet);
                        }
                        else
                        {
                       
                        }
                       
                    }
                }
            }
            catch (Exception ex)
            {

                return Json(new { flag=-2}, JsonRequestBehavior.AllowGet);
            }

            finally
            {
                Dispose();
            }
            return Json(new { flag = -1 }, JsonRequestBehavior.AllowGet);
        }

        [SessionAuthorize]
        public ActionResult EmployeeShiftDetails()
        {
            return View();
        }
        public JsonResult GetShiftType()
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
                        cmd.Parameters.AddWithValue("@ID_Module", 1);
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
            catch(Exception ex)
            {

            }
            finally
            {
                Dispose();
            }
            return Json(new { flag = false }, JsonRequestBehavior.AllowGet);
        }


        public JsonResult GetSingleDateShift(int ColumnId )
        {

            try
            {
                string jResult = string.Empty;
                using (SqlConnection con = new SqlConnection(consString))
                {
                    using (SqlCommand cmd = new SqlCommand("VTwo.USPTblShiftDetailsSelect"))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Connection = con;
                        cmd.Parameters.AddWithValue("@ID_TblShiftdetails", ColumnId);
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
                                return Json(new { result = jResult }, JsonRequestBehavior.AllowGet);
                        }
                        else
                        {

                        }

                    }
                }
            }
            catch (Exception ex)
            {

                return Json(new { flag = -2 }, JsonRequestBehavior.AllowGet);
            }

            finally
            {
                Dispose();
            }
            return Json(new { flag = -1 }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult UpdateEmpShiftDetails(int Id, string OldStartTime, string OldEndTime,int oldShiftType, string newSatrtTime, string newStartDate,int newShiftType,string ShiftTeam)
        {                     
           try
            {
                bool mailFlag = false;
                bool Requestflag = false;
                int MonitoringShiftHours = Convert.ToInt32(ConfigurationManager.AppSettings["MonitoringShiftTime"]);
                int DbaShiftHours = Convert.ToInt32(ConfigurationManager.AppSettings["DbaShiftTime"]);
                DateTimeOffset Startdate = new DateTimeOffset();
                DateTimeOffset Enddate = new DateTimeOffset();
                string newSarttDate, newendtDate;
                TimeSpan _offset = new TimeSpan(5, 30, 00);
                Int32 _hr = 0, _min = 0, _sec = 0;
                string jResult = string.Empty;
                string[] ary = newStartDate.Split('-');
                string[] _timetemp = newSatrtTime.Split(':');
                if (_timetemp.Length == 2)
                {
                    _hr = Convert.ToInt32(_timetemp[0]);
                    _min = Convert.ToInt32(_timetemp[1]);
                }
                if (_timetemp.Length == 3)
                {
                    _hr = Convert.ToInt32(_timetemp[0]);
                    _min = Convert.ToInt32(_timetemp[1]);
                    _sec = Convert.ToInt32(_timetemp[2]);
                }
                Startdate = new DateTimeOffset(Convert.ToInt16(ary[0]), Convert.ToInt16(ary[1]), Convert.ToInt16(ary[2]),
                     _hr, _min, _sec, _offset);

                newSarttDate= Startdate.ToString("yyyy-MM-dd HH:mm:ss zzz");

                if(newShiftType!=2)
                {
                    newendtDate = newSarttDate;
                }
                else
                {
                    if (ShiftTeam == "Monitoring")
                    {
                        Enddate = Startdate.AddHours(MonitoringShiftHours);
                    }
                    else
                    {
                        Enddate = Startdate.AddHours(DbaShiftHours);
                    }
                    newendtDate = Enddate.ToString("yyyy-MM-dd HH:mm:ss zzz");
                }

               
                using (SqlConnection con = new SqlConnection(consString))
                {
                    using (SqlCommand cmd = new SqlCommand("VTwo.USPTblShiftChangeReguestUpdate"))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Connection = con;
                       
                        cmd.Parameters.AddWithValue("@ID_TblShiftChangeReguest", 0);
                        cmd.Parameters.AddWithValue("@SCRExistingShiftStartTime", OldStartTime);
                        cmd.Parameters.AddWithValue("@SCRExistingShiftEndTime", OldEndTime);
                        cmd.Parameters.AddWithValue("@SCRNewShiftStartTime", newSarttDate);
                        cmd.Parameters.AddWithValue("@SCRNewShiftEndTime", newendtDate);
                        cmd.Parameters.AddWithValue("@FK_ShiftTypeNew", newShiftType);
                        cmd.Parameters.AddWithValue("@FK_ShiftTypeExisting", oldShiftType);
                        cmd.Parameters.AddWithValue("@FK_TblShiftDetails", Id);
                        cmd.Parameters.AddWithValue("@FK_EnteredBy", 1);
                        cmd.Parameters.AddWithValue("@FK_ApprovedBY", 1);
                        cmd.Parameters.AddWithValue("@Action", 1);
                        con.Open();
                        Requestflag = Convert.ToBoolean(cmd.ExecuteNonQuery());
                        if (Requestflag == true)
                        {
                            //CommonFunctions functionMail = new CommonFunctions();
                            //string UpdateShiftLink=(ConfigurationManager.AppSettings["UpdateShiftLink"]); 
                            //string msgBody = "<a id='resetLink' href='" + UpdateShiftLink + "'> Approve link </a>"; 
                            //mailFlag = functionMail.SendEmail("sanal@nuvento.com", "", "", "Request For Shift Changing", "Hi, <br/>&nbsp;&nbsp;&nbsp; To approve shift use this link  " +
                            //          msgBody);
                            return Json(new { flag = 1 }, JsonRequestBehavior.AllowGet);
                            
                        }
                        else
                        {
                           // return Json(new { flag = mailFlag }, JsonRequestBehavior.AllowGet);
                        }
                        con.Close();
                        return Json(new { flag = 1 }, JsonRequestBehavior.AllowGet);


                    }
                }
            }
            catch (Exception ex)
            {

                return Json(new { flag = -2 }, JsonRequestBehavior.AllowGet);
            }

            finally
            {
                Dispose();
            }
          //  return Json(new { flag = -1 }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult InsertDoubleShift(string empName, string newSatrtTime, string newStartDate, string newShiftType,int shiftTypeId)
        {
            try
            {
                DataTable dta = new DataTable();
                dta.Columns.Add("Name", typeof(System.String));
                dta.Columns.Add("TeamName", typeof(System.String));
                dta.Columns.Add("StartDate", typeof(System.String));
                dta.Columns.Add("EndDate", typeof(System.String));
                DataRow dtrow;
                bool mailFlag = false;
                bool Requestflag = false;
                int MonitoringShiftHours = Convert.ToInt32(ConfigurationManager.AppSettings["MonitoringShiftTime"]);
                int DbaShiftHours = Convert.ToInt32(ConfigurationManager.AppSettings["DbaShiftTime"]);
                DateTimeOffset Startdate = new DateTimeOffset();
                DateTimeOffset Enddate = new DateTimeOffset();
                string newSarttDate, newendtDate;
                TimeSpan _offset = new TimeSpan(5, 30, 00);
                Int32 _hr = 0, _min = 0, _sec = 0;
                string jResult = string.Empty;
                string[] ary = newStartDate.Split('-');
                string[] _timetemp = newSatrtTime.Split(':');
                if (_timetemp.Length == 2)
                {
                    _hr = Convert.ToInt32(_timetemp[0]);
                    _min = Convert.ToInt32(_timetemp[1]);
                }
                if (_timetemp.Length == 3)
                {
                    _hr = Convert.ToInt32(_timetemp[0]);
                    _min = Convert.ToInt32(_timetemp[1]);
                    _sec = Convert.ToInt32(_timetemp[2]);
                }
                Startdate = new DateTimeOffset(Convert.ToInt16(ary[0]), Convert.ToInt16(ary[1]), Convert.ToInt16(ary[2]),
                     _hr, _min, _sec, _offset);

                newSarttDate = Startdate.ToString("yyyy-MM-dd HH:mm:ss zzz");

                if (shiftTypeId != 2)
                {
                    newendtDate = newSarttDate;
                }
                else
                {
                    if (newShiftType == "Monitoring")
                    {
                        Enddate = Startdate.AddHours(MonitoringShiftHours);
                    }
                    else
                    {
                        Enddate = Startdate.AddHours(DbaShiftHours);
                    }
                    newendtDate = Enddate.ToString("yyyy-MM-dd HH:mm:ss zzz");
                }
                dtrow= dta.NewRow();
                dtrow[0] = empName;
                dtrow[1] = newShiftType;
                dtrow[2] = newSarttDate;
                dtrow[3] = newendtDate;
                if (!dtrow.IsNull(0))
                {
                    dta.Rows.Add(dtrow);
                }

                if (dta.Rows.Count > 0)
                {

                    using (SqlConnection con = new SqlConnection(consString))
                    {
                        using (SqlCommand cmd = new SqlCommand("VTwo.USPTblShiftDetailsInsertDoubleShift"))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Connection = con;
                            cmd.Parameters.AddWithValue("@ShiftDetails", dta);
                            cmd.Parameters.AddWithValue("@FK_ModifiedBy", 1);
                            cmd.Parameters.AddWithValue("@FK_EnteredBy", 1);

                            con.Open();
                            //cmd.ExecuteNonQuery();
                            //SqlDataReader dataReader = cmd.ExecuteReader();
                            DataTable dataTable = new DataTable();
                            dataTable.Load(cmd.ExecuteReader());
                           // int Returnvalue = Convert.ToInt32(dataTable.Rows[0]["Returnvalue"]);
                            con.Close();
                            return Json(new { flag = 1 }, JsonRequestBehavior.AllowGet);

                        }
                    }
                }
            }
            catch (Exception ex)
            {

                return Json(new { flag = -2 }, JsonRequestBehavior.AllowGet);
            }

            finally
            {
                Dispose();
            }
              return Json(new { flag = -1 }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult DeleteEmpShiftDetails(int Id)
        {

            try
            {
                string jResult = string.Empty;
                using (SqlConnection con = new SqlConnection(consString))
                {
                    using (SqlCommand cmd = new SqlCommand("VTwo.USPTblShiftDetailsDelete"))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Connection = con;
                        cmd.Parameters.AddWithValue("@ID_ShiftDetails", Id);
                        con.Open();
                        cmd.ExecuteNonQuery();
                        SqlDataReader dataReader = cmd.ExecuteReader();
                        DataTable dataTable = new DataTable();
                        dataTable.Load(dataReader);
                        con.Close();
                       
                        return Json(new { flag = 1 }, JsonRequestBehavior.AllowGet);
                       

                    }
                }
            }
            catch (Exception ex)
            {

                return Json(new { flag = -2 }, JsonRequestBehavior.AllowGet);
            }

            finally
            {
                Dispose();
            }
           // return Json(new { flag = -1 }, JsonRequestBehavior.AllowGet);
        }
        
    }
}