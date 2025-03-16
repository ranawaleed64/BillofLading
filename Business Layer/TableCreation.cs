using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
namespace BillOfLading
{
    public class TableCreation
    {
        #region TableCreation
        public TableCreation()
        {
            try
            {
                this.billOfLading();
            }
            catch (Exception ex)
            {
                EventHandler.oApplication.StatusBar.SetText("Table Creation Failed: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }

        #region ...Fields Creation
        public void billOfLading()
        {
            try
            {
                this.bolHeader();
                this.BOLDetail();



                if (!GlobalVariables.oGFun.UDOExists("BLC_BOLD"))
                {
                    string[,] FindField = new string[,] { { "DocNum", "DocNum" } };
                    GlobalVariables.oGFun.RegisterUDO("BLC_BOLD", "Bill of Lading", SAPbobsCOM.BoUDOObjType.boud_Document, FindField, "BLC_BOLD", "", "BLC_BLD1");
                    FindField = null;
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
        public void bolHeader()
        {
            try
            {

                GlobalVariables.oGFun.CreateTable("BLC_BOLD", "Bill of Lading Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "VC", "Vendor Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
               // GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "SRCENTRY", "Source Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "VN", "Vendor Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "DS", "Document Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "DATE", "Date", SAPbobsCOM.BoFieldTypes.db_Date, 11);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "BN", "BL Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "PN", "PO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "REMARK", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 200);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "40HQ", "Total 40HQ Container", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "20GP", "Total 20GP Container", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "SONO", "SO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "LandCost", "Total Landed Cost", SAPbobsCOM.BoFieldTypes.db_Float, 11, SAPbobsCOM.BoFldSubTypes.st_Sum);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "Draft", "Draft Number.", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "Status", "Status.", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "PodETA", "POD ETA.", SAPbobsCOM.BoFieldTypes.db_Date, 11);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "FeedETD", "Feeder ETD.", SAPbobsCOM.BoFieldTypes.db_Date, 11);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BOLD", "MainETD", "Main ETD.", SAPbobsCOM.BoFieldTypes.db_Date, 11);

            }
            catch (Exception ex)
            {
                EventHandler.oApplication.MessageBox(ex.Message);
            }
        }
        public void BOLDetail()
        {
            try
            {
                GlobalVariables.oGFun.CreateTable("BLC_BLD1", "Bill of Lading Detail", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);

                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "PODocEntry", "PO DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric,11);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "PONO", "PO Number", SAPbobsCOM.BoFieldTypes.db_Numeric,11);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "POLINE", "PO LineNum", SAPbobsCOM.BoFieldTypes.db_Numeric,11);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "IC", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "IN", "ItemName", SAPbobsCOM.BoFieldTypes.db_Alpha, 200);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "POQ", "PO Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 11, SAPbobsCOM.BoFldSubTypes.st_Quantity);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "Allocated", "Allocated Qty", SAPbobsCOM.BoFieldTypes.db_Float, 11, SAPbobsCOM.BoFldSubTypes.st_Quantity);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "Commodity", "Commodity", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "Model", "Model.", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "ContNo", "Container No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "ContType", "Container Type.", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "IncTerms", "Incoterms Shipping Terms.", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "FeedPort", "Feeder Port.", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "FeedETD", "Feeder ETD.", SAPbobsCOM.BoFieldTypes.db_Date, 11);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "MainPort", "Main Port", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "MainETD", "Main Port ETD", SAPbobsCOM.BoFieldTypes.db_Date, 11);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "Pod", "POD", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "PodETA", "POD ETA", SAPbobsCOM.BoFieldTypes.db_Date, 11);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "GateDate", "Gate Out Date", SAPbobsCOM.BoFieldTypes.db_Date, 11);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "EmptyDate", "Empty in Date", SAPbobsCOM.BoFieldTypes.db_Date, 11);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "CNEE", "CNEE", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "Notify", "Notify.", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "Forwarder", "Forwarder.", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "ShipLine", "Shipping Line.", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "FreeDays", "Free Days.", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "Cost", "O/F Cost", SAPbobsCOM.BoFieldTypes.db_Float, 11,SAPbobsCOM.BoFldSubTypes.st_Sum);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "Telex", "Telex", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
                GlobalVariables.oGFun.CreateUserFields("@BLC_BLD1", "TelexDate", "Telex Date", SAPbobsCOM.BoFieldTypes.db_Date, 11);
 
            }

            catch (Exception ex)
            {
                EventHandler.oApplication.MessageBox(ex.Message);
            }
        }
        #endregion

  
    }

    #endregion
}
