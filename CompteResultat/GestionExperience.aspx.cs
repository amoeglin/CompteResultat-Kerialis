using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;

using CompteResultat.BL;
using CompteResultat.DAL;
using CompteResultat.Common;
using System.Web.DynamicData;
using System.Web.ModelBinding;
using OfficeOpenXml.Style;
using OfficeOpenXml;

namespace CompteResultat
{
    public partial class GestionExperience : System.Web.UI.Page
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                cmdImport.Attributes.Add("onclick", "jQuery('#" + uploadExcel.ClientID + "').click();return false;");

                if (!IsPostBack)
                {
                    lbAssur.SelectedIndex = 0;
                    rptExp.DataBind();
                    //UpdateTreeView(C.cDEFAULTASSUREUR);
                }
                else
                {
                    //Handle the Delete Event
                    if (Request.Form["cmdDelete"] != null) { }

                    //Handle the import of groups & guarantees
                    if (uploadExcel.PostedFile != null)
                    {
                        if (uploadExcel.PostedFile.FileName.Length > 0)
                        {
                            //Import Garanties
                            string excelFile = Path.GetFileName(uploadExcel.PostedFile.FileName);

                            string excelDirectory = Path.Combine(Request.PhysicalApplicationPath, C.excelFolder);
                            string fullUploadPath = Path.Combine(excelDirectory, excelFile);
                            uploadExcel.PostedFile.SaveAs(fullUploadPath);

                            //string compId = lbCompanies.SelectedItem.Value.ToString();
                            //int iCompId = -1;
                            string assurId = lbAssur.SelectedItem.Value.ToString();
                            string assurName = lbAssur.SelectedItem.Text.ToString();
                            int iAssurId = -1;

                            if (int.TryParse(assurId, out iAssurId))
                            {
                                if (iAssurId == -1)
                                    assurName = C.cDEFAULTASSUREUR;
                                BLCadencier.ImportCadencierForAssureur(assurName, fullUploadPath, true);

                                //refresh the tree
                                lbAssur.DataBind();

                                if (ItemExists(assurName))
                                {
                                    SelectItem(assurName);
                                    rptExp.DataBind();
                                    //UpdateTreeView(assurName);
                                }
                                else
                                {
                                    if (lbAssur.Items.Count > 0)
                                    {
                                        SelectItem(lbAssur.Items[0].Text);
                                        rptExp.DataBind();
                                        //UpdateTreeView(lbAssur.Items[0].Text);
                                    }
                                    else
                                        rptExp.DataBind();
                                    //    tvGaranties.Nodes.Clear();
                                }


                                //lbAssur.DataBind();
                                //string selAssur = Session["SelectedAssureurName"].ToString();
                                //SelectItem(selAssur);
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { UICommon.HandlePageError(ex, this.Page, "GestionExperience::Page_Load"); }
        }

        public IEnumerable<C_TempExpData> GetExperience()
        {
            string assurName;

            try
            {
                if (lbAssur.SelectedItem != null)
                    assurName = lbAssur.SelectedItem.Text.ToString();
                else
                    assurName = C.cDEFAULTASSUREUR;

                return C_TempExpData.GetExpData(2018);
            }
            catch (Exception ex) { UICommon.HandlePageError(ex, this.Page, "GestionExperience::GetExperience"); return null; }
        }

        public List<Assureur> GetAssureurs([Control] bool chkAssur)
        {
            try
            {
                List<Assureur> assur;

                if (chkAssur)
                {
                    assur = Assureur.GetAssureursWithoutCadencier();
                }
                else
                {
                    assur = Assureur.GetAllAssureurs();
                    assur.Insert(0, new Assureur { Id = -1, Name = C.cDEFAULTASSUREUR });
                }

                return assur;
            }
            catch (Exception ex) { UICommon.HandlePageError(ex, this.Page, "GestionExperience::GetAssureurs"); return null; }
        }        

        protected void lbAssur_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (lbAssur.SelectedItem != null)
                {
                    Session["SelectedAssureurIndex"] = lbAssur.SelectedIndex;
                    Session["SelectedAssureurName"] = lbAssur.SelectedItem.Text;

                    string assurId = lbAssur.SelectedItem.Value.ToString();
                    string assurName = lbAssur.SelectedItem.Text.ToString();

                    rptExp.DataBind();
                    //UpdateTreeView(assurName);
                }
                else
                {
                    if (lbAssur.Items.Count > 0)
                        lbAssur.Items[0].Selected = true;
                }
            }
            catch (Exception ex) { UICommon.HandlePageError(ex, this.Page, "GestionExperience::lbAssur_SelectedIndexChanged"); }
        }

        protected void cmdExport_Click(object sender, EventArgs e)
        {
            string uploadPath = "";

            try
            {
                string assurName = lbAssur.SelectedItem.Text.ToString();

                List<Cadencier> groupsSante = Cadencier.GetCadencierForAssureur(assurName); 
                ExcelPackage pack = BLCadencier.ExportCadencierForAssureur(assurName);

                uploadPath = Path.Combine(Request.PhysicalApplicationPath, C.uploadFolder);
                uploadPath = Path.Combine(uploadPath, User.Identity.Name + "_" + assurName + "_Cadencier.xlsx");

                pack.SaveAs(new FileInfo(uploadPath));

                UICommon.DownloadFile(this.Page, uploadPath);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Thread was being aborted.")
                    log.Error(ex.Message);

                var myCustomValidator = new CustomValidator();
                myCustomValidator.IsValid = false;
                myCustomValidator.ErrorMessage = ex.Message;
                Page.Validators.Add(myCustomValidator);
            }
        }

        protected void cmdDelete_Click(object sender, EventArgs e)
        {
            try
            {
                //handle the delete event
                string assurName = lbAssur.SelectedItem.Text.ToString();

                Session["SelectedAssureurIndex"] = lbAssur.SelectedIndex;

                if (!string.IsNullOrWhiteSpace(assurName))
                    Cadencier.DeleteCadencierWithSpecificAssureurName(assurName);
                else
                    throw new Exception("Please select the name of the 'Assureur' for which you want to delete the Cadencier!");

                //refresh the tree
                lbAssur.DataBind();

                if (ItemExists(assurName))
                {
                    SelectItem(assurName);
                    rptExp.DataBind();
                    //UpdateTreeView(assurName);
                }
                else
                {
                    if (lbAssur.Items.Count > 0)
                    {
                        SelectItem(lbAssur.Items[0].Text);
                        rptExp.DataBind();
                        //UpdateTreeView(lbAssur.Items[0].Text);
                    }
                    else
                        //tvGaranties.Nodes.Clear();
                        rptExp.DataBind();
                }
            }
            catch (Exception ex) { UICommon.HandlePageError(ex, this.Page, "GestionExperience::cmdDelete_Click"); }
        }

        protected void lbAssur_DataBound(object sender, EventArgs e)
        {
            try
            {
                //if no item is selected, select the first item in the list
                if (lbAssur.SelectedIndex == -1 && lbAssur.Items.Count > 0)
                {
                    SelectItem(lbAssur.Items[0].Text);
                    rptExp.DataBind();                    
                    
                    //UpdateTreeView(lbAssur.Items[0].Text);
                }

                EnableDisableButtons();
            }
            catch (Exception ex) { UICommon.HandlePageError(ex, this.Page, "GestionExperience::lbAssur_DataBound"); }
        }

        protected void rptExp_ItemDataBound(object sender, RepeaterItemEventArgs e)
        {
            Repeater rpt = sender as Repeater; // Get the Repeater control object.
            //PlaceHolder phHeader;

            //if (e.Item.ItemType == ListItemType.Header)
            //{
            //    //phHeader
            //    phHeader = e.Item.FindControl("phHeader") as PlaceHolder;
            //    if (phHeader != null)
            //    {
            //        phHeader.Visible = true;
            //    }
            //}

            // If the Repeater contains no data.
            if (rpt != null )
            {
                if (e.Item.ItemType == ListItemType.Footer)
                {
                    // Show info Label (if no data is present).
                    //Label lblNoData = e.Item.FindControl("lblEmpty") as Label;
                    //if (lblNoData != null)
                    //{
                    //    lblNoData.Visible = true;
                    //}

                    if (rpt.Items.Count < 1)
                    {
                        rptExp.Visible = false;
                        phHeader.Visible = true;
                    }
                    else
                    {
                        rptExp.Visible = true;
                        phHeader.Visible = false;
                    }
                }

                //if (e.Item.ItemType == ListItemType.Header)
                //{
                //    //phHeader
                //    PlaceHolder phHeader = e.Item.FindControl("phHeader") as PlaceHolder;
                //    if (phHeader != null)
                //    {
                //        phHeader.Visible = false;
                //    }
                //}
            }  
        }


        #region Private Methods

        private void EnableDisableButtons()
        {
            try
            {
                if (lbAssur.SelectedItem == null)
                {
                    cmdExport.Enabled = false;
                    cmdImport.Enabled = false;
                    cmdDelete.Enabled = false;
                }
                else
                {
                    cmdExport.Enabled = true;
                    cmdImport.Enabled = true;
                    cmdDelete.Enabled = true;
                }
            }
            catch (Exception ex) { UICommon.HandlePageError(ex, this.Page, "GestionExperience::EnableDisableButtons"); }
        }

        private void SelectItem(string itemName)
        {
            try
            {
                foreach (ListItem li in lbAssur.Items)
                {
                    if (li.Text == itemName)
                        li.Selected = true;
                    else
                        li.Selected = false;
                }
            }
            catch (Exception ex) { UICommon.HandlePageError(ex, this.Page, "GestionExperience::SelectItem"); }
        }

        private bool ItemExists(string itemName)
        {
            try
            {
                foreach (ListItem li in lbAssur.Items)
                {
                    if (li.Text == itemName)
                        return true;
                }

                return false;
            }
            catch (Exception ex) { UICommon.HandlePageError(ex, this.Page, "GestionExperience::ItemExists"); return false; }
        }




        #endregion

        protected void cmdRecreate_Click(object sender, EventArgs e)
        {
            try
            {
                BLGroupsAndGaranties.RecreateGroupsGarantiesSanteFromPresta();
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);

                var myCustomValidator = new CustomValidator();
                myCustomValidator.IsValid = false;
                myCustomValidator.ErrorMessage = ex.Message;
                Page.Validators.Add(myCustomValidator);
            }
        }
    }
}