using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace BillOfLading
{
internal class GRPO
    {
        #region Variables
        SAPbouiCOM.Form oFrom;
        SAPbouiCOM.Matrix oMatrix;
        SAPbobsCOM.Recordset oRecordSet;
        SAPbouiCOM.Form frmPo, frmGRPO, frmUDF;
        SAPbobsCOM.Company targetCompany;
        SAPbobsCOM.GeneralService oGeneralService = null;
        SAPbobsCOM.GeneralData oGeneralData = null;
        SAPbobsCOM.GeneralDataCollection oSons = null;
        SAPbobsCOM.GeneralData oSon = null;
        SAPbobsCOM.CompanyService sCmp = null;
        SAPbobsCOM.Recordset oRecordSet1;
        string query = "";
        XmlDocument doc = new XmlDocument();
        int branch = 0, InvoiceNumber = 0;
        SAPbouiCOM.Item postIRBMItem, btnResend;
        SAPbouiCOM.DBDataSource oDBHeader;
        #endregion

        public void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.EventType)
                {

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        try
                        {
                            switch (pVal.ItemUID)
                            {
                                case "btnCopy":
                                    if (pVal.BeforeAction == false)
                                    {
                                        oFrom = EventHandler.oApplication.Forms.Item(FormUID);
                                        GlobalVariables.frmAP = oFrom;
                                        SAPbouiCOM.EditText VendorCode = oFrom.Items.Item("4").Specific;

                                        if (string.IsNullOrWhiteSpace(VendorCode.Value))
                                        {
                                            EventHandler.oApplication.StatusBar.SetText("Customer Code is Missing.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                            return;

                                        }

                                        var oObj = new BillofLadingWizard();
                                        oObj.LoadblWizard(VendorCode.Value);
                                    }
                                    break;
                                //case "btnResend":
                                //    if (pVal.BeforeAction == false)
                                //    {
                                //        oFrom = EventHandler.oApplication.Forms.Item(FormUID);
                                //        GlobalVariables.frmAP = oFrom;

                                //        try
                                //        {
                                //            //ProcessInvoiceAndReserveResend(FormUID, out string errorMessage);

                                //            //if (!string.IsNullOrEmpty(errorMessage))
                                //            //{
                                //            //    GlobalVariables.oGFun.SendMessage("AR to AP Failure Notification", new string[] { GlobalVariables.oCompany.UserName }, 13, Convert.ToInt32(oDBHeader.GetValue("DocEntry", 0)), errorMessage);
                                //            //    EventHandler.oApplication.StatusBar.SetText($"Error: {errorMessage}",
                                //            //        SAPbouiCOM.BoMessageTime.bmt_Short,
                                //            //        SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                //            //}
                                //            //else
                                //            //{
                                //            //    oFrom = EventHandler.oApplication.Forms.Item(FormUID);
                                //            //    oDBHeader = oFrom.DataSources.DBDataSources.Item("OINV");
                                //            //    GlobalVariables.oGFun.SendMessage("AR to AP Succes Notification", new string[] { GlobalVariables.oCompany.UserName }, 13, Convert.ToInt32(oDBHeader.GetValue("DocEntry", 0)), "AP Reserve Invoice Posted");

                                //            //}
                                //        }
                                //        catch (Exception ex)
                                //        {
                                //            EventHandler.oApplication.StatusBar.SetText($"Exception: {ex.Message}",
                                //                SAPbouiCOM.BoMessageTime.bmt_Short,
                                //                SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                //        }




                                //    }
                                //    break;

                            }
                        }
                        catch (Exception ex)
                        {
                            EventHandler.oApplication.StatusBar.SetText("Item Pressed Event Failed" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        finally
                        {
                        }

                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        try
                        {
                            if (pVal.BeforeAction == false)
                            {

                                oFrom = EventHandler.oApplication.Forms.Item(FormUID);
                                try
                                {
                                    postIRBMItem = oFrom.Items.Item("btnCopy");
                                   // btnResend = oFrom.Items.Item("btnResend");
                                }
                                catch
                                {
                                    postIRBMItem = null;
                                   // btnResend = null;
                                }

                                if (postIRBMItem == null)
                                {
                                    GlobalVariables.oGFun.AddControls(oFrom, "btnCopy", "33", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "DOWN", 1, 1, "", "Copy from Bill of Lading", 140);
                                    //GlobalVariables.oGFun.AddControls(oFrom, "btnResend", "34", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "Down", 1, 1, "", "Resend", 80);

                                    postIRBMItem = oFrom.Items.Item("btnCopy");
                                   // btnResend = oFrom.Items.Item("btnResend");
                                    //btnResend.Left = btnResend.Left + 30;
                                    postIRBMItem.Visible = true;
                                    //btnResend.Visible = false;
                                }

                            }
                        }
                        catch (Exception ex)
                        {
                            EventHandler.oApplication.StatusBar.SetText("Item Pressed Event Failed" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        finally
                        {
                        }

                        break;

                }
            }
            catch (Exception ex)
            {
                EventHandler.oApplication.MessageBox(ex.Message);
            }
            finally
            {
            }
        }
        public void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                oFrom = EventHandler.oApplication.Forms.ActiveForm;

                switch (pVal.MenuUID)
                {


                    case "1282":
                        {
                            postIRBMItem = oFrom.Items.Item("btnCopy");
                           // btnResend = oFrom.Items.Item("btnResend");
                            if (oFrom.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                postIRBMItem.Visible = true;
                                //btnResend.Visible = false;

                            }
                            else
                            {
                                //btnResend.Visible = true;
                                postIRBMItem.Visible = false;

                            }
                            break;
                        }

                }
            }
            catch (Exception ex)
            {
                EventHandler.oApplication.MessageBox(ex.Message);
            }
        }

        public bool PopulateMultiplePODetails(List<CopyData> dt)
        {
            SAPbouiCOM.Form aForm = GlobalVariables.frmAP;
            SAPbouiCOM.Form frmUDF = EventHandler.oApplication.Forms.Item(GlobalVariables.frmAP.UDFFormUID);

            try
            {
                aForm.Freeze(true);
                aForm.Select();

                oMatrix = aForm.Items.Item("38").Specific;
                
                int row = 1;
                for (int introw = 0; introw < dt.Count; introw++)
                {
                    //BLData data = new BLData();

                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(row).Specific).Value = dt[introw].ItemCode;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(row).Specific).Value = Convert.ToString(dt[introw].Qty);
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_BLNUM").Cells.Item(row).Specific).Value = Convert.ToString(dt[introw].BLNo);
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("14").Cells.Item(row).Specific).Value = Convert.ToString(dt[introw].UnitPrice);
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("163").Cells.Item(row).Specific).Value = Convert.ToString(dt[introw].containerNo);
                    //string containers = string.Join(",", dt[introw].containerNo);

                    //((SAPbouiCOM.EditText)oMatrix.Columns.Item("163").Cells.Item(row).Specific).Value = containers;



                    row = row + 1;
                }
                aForm.Freeze(false);
                return true;
            }
            catch (Exception ex)
            {
                EventHandler.oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                aForm.Freeze(false);
                return false;
            }
        }
    }
}
