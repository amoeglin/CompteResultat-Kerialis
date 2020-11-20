using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.IO;
using System.Data;

using System.Data.Entity.Migrations;
using System.Data.Entity.Validation;
using System.Text.RegularExpressions;
using System.Globalization;

using Excel;
using CompteResultat.DAL;
using CompteResultat.Common;
using OfficeOpenXml;
using OfficeOpenXml.Style;


namespace CompteResultat.BL
{
    public class BLCadencier
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


        public static void ImportCadencierForAssureur(string assureurName, string excelFilePath, bool firstRowAsColumnNames)
        {
            try
            {
                int year;
                DateTime debutSurv;
                DateTime finSurv;
                int month;
                double cumul;

                //read Excel file into datatable
                DataTable dt = G.ExcelToDataTable(excelFilePath, firstRowAsColumnNames);

                // delete all rows in DB Tables with the specified assurName 
                Cadencier.DeleteCadencierWithSpecificAssureurName(assureurName);

                foreach (DataRow row in dt.Rows)
                {
                    //### validate => all fields must be specified                    
                    //codeActe = row[C.eExcelGroupsGaranties.CodeActe.ToString()].ToString();

                    if (!Int32.TryParse(row[C.eExcelCadencier.Year.ToString()].ToString(), out year))
                        throw new Exception("One of the provided 'Year' values is not valid for the Cadencier you are trying to import !");

                    if (!Int32.TryParse(row[C.eExcelCadencier.Month.ToString()].ToString(), out month))
                        throw new Exception("One of the provided 'Month' values is not valid for the Cadencier you are trying to import !");

                    if (!double.TryParse(row[C.eExcelCadencier.Cumul.ToString()].ToString(), out cumul))
                        throw new Exception("One of the provided 'Cumul' values is not valid for the Cadencier you are trying to import !");

                    if (!DateTime.TryParse(row[C.eExcelCadencier.DebutSurvenance.ToString()].ToString(), out debutSurv))
                        throw new Exception("One of the provided 'DebutSurvenance' values is not valid for the Cadencier you are trying to import !");

                    if (!DateTime.TryParse(row[C.eExcelCadencier.FinSurvenance.ToString()].ToString(), out finSurv))
                        throw new Exception("One of the provided 'FinSurvenance' values is not valid for the Cadencier you are trying to import !");

                    int id = Cadencier.InsertCadencier(new Cadencier
                    {
                        AssureurName = assureurName,
                        Year=year,
                        DebutSurvenance=debutSurv,
                        FinSurvenance=finSurv,
                        Month=month,
                        Cumul=cumul                        
                    });
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                throw ex;
            }
        }

        public static ExcelPackage ExportCadencierForAssureur(string assureurName)
        {
            try
            {
                List<Cadencier> cad = Cadencier.GetCadencierForAssureur(assureurName);

                ExcelPackage pck = new ExcelPackage();
                var ws = pck.Workbook.Worksheets.Add(assureurName);

                //write the header
                //ws.Cells["A3"].Style.Numberformat.Format = "yyyy-mm-dd";
                //ws.Column(2).Style.Numberformat.Format = "dd-mm-yyyy";
                //ws.Column(3).Style.Numberformat.Format = "dd-mm-yyyy";
                ws.Column(2).Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                ws.Column(3).Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                ws.Cells[1, 1].Value = "Year";
                ws.Cells[1, 2].Value = "DebutSurvenance";
                ws.Cells[1, 3].Value = "FinSurvenance";
                ws.Cells[1, 4].Value = "Month";
                ws.Cells[1, 5].Value = "Cumul";

                int row = 2;

                foreach (Cadencier c in cad)
                {
                    ws.Cells[row, 1].Value = c.Year;
                    ws.Cells[row, 2].Value = c.DebutSurvenance;
                    ws.Cells[row, 3].Value = c.FinSurvenance;
                    ws.Cells[row, 4].Value = c.Month;
                    ws.Cells[row, 5].Value = c.Cumul;

                    row++;
                }

                return pck;
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                throw ex;
            }
        }


    }
}
