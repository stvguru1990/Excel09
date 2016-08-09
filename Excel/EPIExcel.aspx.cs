using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Collections.Specialized;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;



namespace Excel
{
    public partial class EPIExcel : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void ExcelButton_Click(object sender, EventArgs e)
        {
            String strConnString = ConfigurationManager.ConnectionStrings["conString"].ConnectionString;
            SqlConnection con = new SqlConnection(strConnString);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "GetAllEmployeeDetails";
            cmd.Connection = con;
            try
            {
                DataTable dtTask = new DataTable();
                SqlCommand command = new SqlCommand("SPExcelUserDetails", con);
                command.CommandType = CommandType.StoredProcedure;
                con.Open();
                SqlDataReader reader = command.ExecuteReader();
                dtTask.Load(reader);
                DumpExcel(dtTask);
                con.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        private void DumpExcel(DataTable dtTask)
        {
            using (ExcelPackage pck = new ExcelPackage())
            {
                //Create the worksheet

               // Response.ClearContent();
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Demo");

                //Load the datatable into the sheet, starting from cell A1. Print the column names on row 1
                ws.Cells["A1"].LoadFromDataTable(dtTask, true);
                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                ws.View.FreezePanes(2,1);

                //Format the header for column 1-3
                //using (ExcelRange rng = ws.Cells["A1:E1"])
                //{
                //    rng.Style.Font.Bold = true;
                //    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Set Pattern for the background to Solid
                //    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));  //Set color to dark blue
                //    rng.Style.Font.Color.SetColor(Color.White);
                //}

                //Example how to Format Column 1 as numeric
                //using (ExcelRange col = ws.Cells[2, 1, 2 + dtTask.Rows.Count, 1])
                //{
                //    col.Style.Numberformat.Format = "#,##0.00";
                //    col.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //}





                //var validation = ws.DataValidations.AddListValidation("$A$10");
                //validation.ShowErrorMessage = true;
                //validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                //validation.ErrorTitle = "Error";
                //validation.Error = "Error Text";
                //sheet with a name : DropDownLists
                //from DropDownLists sheet, get values from cells: !$A$1:$A$10
                //var formula = "=DropDownLists!$A$1:$A$10";
                //validation.Formula.ExcelFormula = formula;




                //Write it back to the client
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=ExcelDemo.xlsx");
                Response.BinaryWrite(pck.GetAsByteArray());   
                Response.End();

            }
        }
    }
}