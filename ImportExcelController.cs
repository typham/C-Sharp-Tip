using Data.EF;
using Data.Models;
using Data.Repositories;
using System;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using System.Web.Configuration;
using System.Data.OleDb;
using System.Data;


namespace SalesHut.Areas.Admin.Controllers
{
    public class ImportController : Controller
    {
        // GET: Admin/Import
        Account _userLogged = Untils.userLogged();
        InvoiceRepository _Invoice = new InvoiceRepository();
        CustomerRepository _Customer = new CustomerRepository();

        [HttpPost]
        public ActionResult Customer(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
                try
                {
                    string path = Path.Combine(Server.MapPath("~/Upload/Excels"),
                    Path.GetFileName(file.FileName));
                    file.SaveAs(path);

                    string excelFilePath = Server.MapPath("~/Upload/Excels/" + file.FileName);
                    string sexcelconnectionstring = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0\"", excelFilePath);

                    OleDbDataAdapter dataAdapter = new OleDbDataAdapter("SELECT * FROM[Sheet1$]", sexcelconnectionstring);
                    DataSet myDataSet = new DataSet();

                    dataAdapter.Fill(myDataSet, "ExcelInfo");

                    System.Data.DataTable dataTable = myDataSet.Tables["ExcelInfo"];

                    foreach (var r in dataTable.AsEnumerable())
                    {
                        var CustomerCode = r["CustomerCode"].ToString();
                        var ExistCustomer = _Customer.SelectByCompany().FirstOrDefault(i => i.CustomerCode == CustomerCode);
                        if (ExistCustomer == null)
                        {
                            var customer = new Customer
                            {
                                CustomerCode = CustomerCode,
                                CustomerName = r["CustomerName"].ToString(),
                                Address = r["Address"].ToString(),
                                City = r["City"].ToString(),
                                Country = r["Country"].ToString(),
                                CustomerSignature = "",
                                Latitude = r["Latitude"].ToString(),
                                Longtitude = r["Longtitude"].ToString(),
                                AddressLatLong = r["Address"].ToString(),
                                CompanyId = _userLogged.CompanyId
                            };

                            _Customer.Insert(customer);
                            _Customer.Save();
                        }
                    }

                    return Redirect(WebConfigurationManager.AppSettings["root"] + "/admin#/customer");
                }
                catch (Exception ex)
                {
                    return Json("ERROR:" + ex.Message.ToString(), JsonRequestBehavior.AllowGet);
                }
            else
            {
                return Json("You have not specified a file.", JsonRequestBehavior.AllowGet);
            }
        }
    }
}