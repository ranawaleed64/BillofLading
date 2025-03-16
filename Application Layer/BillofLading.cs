using Microsoft.SqlServer.Server;
using Microsoft.VisualBasic;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design.Serialization;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Header;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView;

namespace BillOfLading
{
    internal class billoflanding
    {
        #region Variables
        SAPbouiCOM.Form frmBld;
        SAPbouiCOM.DBDataSource oDBDSHeader;
        SAPbouiCOM.DBDataSource oDBDSDetail;
        SAPbouiCOM.Matrix oMatrix, matrix;
        SAPbobsCOM.Recordset oRecordSet;
        SAPbouiCOM.Form frmPo, frmApReserve, frmUDF;
        SAPbouiCOM.ComboBox DocType;
        SAPbouiCOM.Button btnPost, btnCopy;
        SAPbouiCOM.EditText oDraft ;
        int RowtoDelete = 0, intSelectedMatrixrow = 0;
        #endregion

        #region LoadForm
        public void loadBillofLading()
        {
            try
            {

                GlobalVariables.oGFun.LoadXML(frmBld, GlobalVariables.bldID, GlobalVariables.bldXML);
                frmBld = EventHandler.oApplication.Forms.Item(GlobalVariables.bldID);
                oDBDSHeader = frmBld.DataSources.DBDataSources.Item("@BLC_BOLD");
                oDBDSDetail = frmBld.DataSources.DBDataSources.Item("@BLC_BLD1");
                oMatrix = frmBld.Items.Item("matrix").Specific;
                this.InitForm();

            }
            catch (Exception ex)
            {
                EventHandler.oApplication.StatusBar.SetText("Load : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        public void InitForm()
        {
            try
            {
                frmBld.Freeze(true);
                if( frmBld.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
                    oDBDSHeader.SetValue("DocNum", 0, GlobalVariables.oGFun.GetCodeGeneration("@BLC_BOLD").ToString());
                    oDBDSHeader.SetValue("CreateDate", 0, GlobalVariables.oGFun.GetServerDate());
                    GlobalVariables.oGFun.SetNewLine(oMatrix, oDBDSDetail);
                    oDBDSHeader.SetValue("U_Status", 0, "Open");

                }

                oMatrix.CommonSetting.EnableArrowKey=true;      
                this.DefineModesForFields();
                oMatrix.AutoResizeColumns();
                frmBld.Freeze(false);

            }
            catch (Exception ex)
            {
                EventHandler.oApplication.MessageBox(ex.Message);
                frmBld.Freeze(false);
            }
            finally
            {

            }
        }

        private void DefineModesForFields()
        {
            try
            {
                frmBld.Items.Item("Item_10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                frmBld.Items.Item("Item_0").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                frmBld.Items.Item("Item_13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                frmBld.Items.Item("Item_14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            catch (Exception ex)
            {
                EventHandler.oApplication.StatusBar.SetText("DefineModesForFields Method Failed: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            finally
            {
            }
        }
        #endregion

   


        #region Validation
        public bool ValidateAll()
        {
            bool functionReturnValue = true;
            try
            {
                if (frmBld.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || frmBld.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    if (string.IsNullOrWhiteSpace(GlobalVariables.oGFun.GetEditTextValue(frmBld, "Item_3")))
                    {
                        EventHandler.oApplication.StatusBar.SetText("Vendor Code Is missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        frmBld.Items.Item("Item_3").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        return false;
                    }


                    else
                    {
                        functionReturnValue = true;

                    }

                }
            }
            catch (Exception ex)
            {
                EventHandler.oApplication.MessageBox(ex.Message);
                functionReturnValue = false;
            }
            finally
            {
            }
            return functionReturnValue;
        }
        #endregion
        public void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                        try
                        {
                            switch (pVal.ItemUID)
                            {
                               
                                case "matrix":
                                    switch (pVal.ColUID)
                                    {
                                        case "Col_3":
                                            if(pVal.BeforeAction==false)
                                            {
                                                frmBld.Freeze(true);

                                                frmBld.Freeze(false);

                                            }

                                            break;
                                    }

                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            frmBld.Freeze(false);
                            EventHandler.oApplication.StatusBar.SetText("Lost Focus Event Failed" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                        }
                        finally
                        {
                        }

                        break;
                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                        try
                        {
                            switch (pVal.ItemUID)
                            {
                             

                                case "matrix":
                                    switch (pVal.ColUID)
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

                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        try
                        {
                            switch (pVal.ItemUID)
                            {
                                case "1":
                                    if (pVal.BeforeAction == true & (frmBld.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | frmBld.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                                    {
                                        if (this.ValidateAll() == false)
                                        {
                                            System.Media.SystemSounds.Asterisk.Play();
                                            BubbleEvent = false;
                                            return;
                                        }
                                    }
                                    break;
                              
                                case "matrix":
                                    switch (pVal.ColUID)
                                    {
                                        case "#":
                                            this.RowtoDelete = pVal.Row;
                                            this.intSelectedMatrixrow = pVal.Row;
                                            break;
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

                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        try
                        {
                            switch (pVal.ItemUID)
                            {
                               
                          
                                    
                            }
                        }
                        catch (Exception ex) 
                        {
                            EventHandler.oApplication.StatusBar.SetText("Item Pressed Event Failed"+ex.Message,SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        finally
                        {
                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        try
                        {
        
                             SAPbouiCOM.DataTable oDataTable = null;
                            SAPbouiCOM.ChooseFromListEvent oCFLE = (SAPbouiCOM.ChooseFromListEvent)pVal;
                            oDataTable = oCFLE.SelectedObjects;
                            if ((oDataTable != null) & pVal.BeforeAction == false && frmBld.Mode!=BoFormMode.fm_FIND_MODE)
                            {
                                switch (pVal.ItemUID)
                                {
                                    case "Item_3":
                                        if (pVal.BeforeAction == false)
                                        {
                                            oDBDSHeader.SetValue("U_VC", 0, oDataTable.GetValue("CardCode", 0));
                                            oDBDSHeader.SetValue("U_VN", 0, oDataTable.GetValue("CardName", 0));
                                        }
                                        break;



                                    case "matrix":
                                        switch (pVal.ColUID)
                                        {
                                            case "item":
                                                {if(pVal.BeforeAction==false)
                                                    {
                                                        oDBDSDetail.SetValue("U_IC", pVal.Row - 1, oDataTable.GetValue("ItemCode", 0));
                                                        oDBDSDetail.SetValue("U_IN", pVal.Row - 1, oDataTable.GetValue("ItemName", 0));
                                                        oMatrix.LoadFromDataSource();

                                                        GlobalVariables.oGFun.SetNewLine(oMatrix, oDBDSDetail,pVal.Row,"item");
                                                    }
                                                }
                                                    
                                                break;

                                        }

                                        break;
                                }
                                }
                        }
                        catch (Exception ex)
                        {
                            EventHandler.oApplication.StatusBar.SetText("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                        }
                        finally
                        {
                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                        try
                        {
                            switch (pVal.ItemUID)
                            {
                                case "matrix":
                                    switch (pVal.ColUID)
                                    {
                                        case "item":
                                            if (pVal.BeforeAction == false)
                                            {
                                              
                                             


       
                                            }
                                            break;
                                    }
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            EventHandler.oApplication.MessageBox(ex.Message);
                            frmBld.Freeze(false);
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
                switch (pVal.MenuUID)
                {
                    case "1281":
                        {
                            break;
                        }
                    case "1282":
                        {
                            this.InitForm();
                        }
                        break;
                    case "1288":
                    case "1289":
                    case "1290":
                    case "1291":
                    case "1304":
                        {

                            this.DefineModesForFields();
                            break;
                        }

                    case "1292":
                        {
                            GlobalVariables.oGFun.SetNewLine(oMatrix, oDBDSDetail);

                            break;
                        }
                    case "1293":
                        {
                            RefereshDeleteRow();
                            BubbleEvent = false;
                            break;
                        }
                    case "1284":
                        {
                           if(pVal.BeforeAction==true)
                            {
                            }
                            break;
                        }

                }
            }
            catch (Exception ex)
            {
                frmBld.Freeze(false);
                EventHandler.oApplication.MessageBox(ex.Message);
            }
        }

  

        #region FormEvents
        public void FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                switch (BusinessObjectInfo.EventType)
                {

                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                        {
                            try
                            {
                                if (BusinessObjectInfo.BeforeAction)
                                {
                                    if (ValidateAll() == false)
                                    {
                                        BubbleEvent = false;
                                        return;
                                    }
                                    else
                                    {
                                    
                                    }
                                }
                                   
                                        
                            }
                            catch (Exception ex)
                            {
                                EventHandler.oApplication.MessageBox(ex.Message);
                                BubbleEvent = false;
                            }
                            finally
                            {
                            }

                            break;
                        }
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                        try
                        {
                            if (BusinessObjectInfo.BeforeAction == false)
                            {
                        
                            }
                        }
                        catch (Exception ex) { }
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

     
        #endregion
        private void RefereshDeleteRow()
        {
            try
            {
                frmBld.Freeze(true);
                RowtoDelete = intSelectedMatrixrow;
                oMatrix.FlushToDataSource();
                oDBDSDetail.RemoveRecord(RowtoDelete-1);
                oMatrix.LoadFromDataSource();
                oMatrix.FlushToDataSource();

                for (int count = 1; count <= oDBDSDetail.Size ; count++)
                {
                    oDBDSDetail.SetValue("LineId", count - 1, count.ToString());

                }

                oMatrix.LoadFromDataSource();
                if (frmBld.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    frmBld.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                frmBld.Freeze(false);
            }
            catch (Exception ex)
            {
                frmBld.Freeze(false);
                throw ex;
            }
        }

    }

}
