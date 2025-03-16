using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using System.Xml;

namespace BillOfLading
{

    //' <summary>
    //' SAP has set of different events For access the controls.
    //' In this module particularly using to control events.
    //' 1) Menu Event using for while the User choose the menus to select the patricular form 
    //' 2) Item Event using for to pass the event Function while user doing process
    //' 3) Form Data Event Using to Insert,Update,Delete data on Date Base 
    //' 4) Status Bar Event will be call when display message to user, message may be will come 
    //'    Warring or Error
    //' </summary>
    //' <remarks></remarks>

    static class EventHandler
    {

        #region " ... Common Variables For SAP ..."
        private static SAPbouiCOM.Application withEventsField_oApplication;
        public static SAPbouiCOM.Application SBO_Application;
        public static bool oFlagCFL = false;
        public static SAPbouiCOM.Application oApplication
        {

            get { return withEventsField_oApplication; }
            //{ return oApplication; }
            set
            {
                if (withEventsField_oApplication != null)
                {
                    //  withEventsField_oApplication.LayoutKeyEvent -= oApplication_LayoutKeyEvent;
                    withEventsField_oApplication.MenuEvent -= oApplication_MenuEvent;
                    withEventsField_oApplication.AppEvent -= oApplication_AppEvent;
                    withEventsField_oApplication.ItemEvent -= oApplication_ItemEvent;
                    withEventsField_oApplication.FormDataEvent -= oApplication_FormDataEvent;
                    withEventsField_oApplication.StatusBarEvent -= oApplication_StatusBarEvent;
                    withEventsField_oApplication.RightClickEvent -= oApplication_RightClickEvent;
                }
                withEventsField_oApplication = value;
                if (withEventsField_oApplication != null)
                {
                    // withEventsField_oApplication.LayoutKeyEvent += oApplication_LayoutKeyEvent;
                    withEventsField_oApplication.MenuEvent += oApplication_MenuEvent;
                    withEventsField_oApplication.AppEvent += oApplication_AppEvent;
                    withEventsField_oApplication.ItemEvent += oApplication_ItemEvent;
                    withEventsField_oApplication.FormDataEvent += oApplication_FormDataEvent;
                    withEventsField_oApplication.StatusBarEvent += oApplication_StatusBarEvent;
                    withEventsField_oApplication.RightClickEvent += oApplication_RightClickEvent;
                }
            }




        }
        #endregion

        #region " ... 1) Menu Event ..."
        private static void oApplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if ((pVal.BeforeAction == true))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1293":
                        case "1284":
                            SAPbouiCOM.Form oMenuForm = EventHandler.oApplication.Forms.ActiveForm;

                            if (oMenuForm.UniqueID == GlobalVariables.bldID)
                            {
                                GlobalVariables.oBilloflanding.MenuEvent(ref pVal, out BubbleEvent);
                                break;
                            }
                            if (oMenuForm.TypeEx == "143")
                            {
                                GlobalVariables.ARinvoice.MenuEvent(ref pVal, out BubbleEvent);
                                break;
                            }

                            break;
                    }
                }

                if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1282":
                        case "1281":
                        case "1287":
                        case "1292":
                        case "1293":
                        case "1294":
                        case "1284":
                        case "519":
                        case "1286":
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                        case "1304":
                            SAPbouiCOM.Form oMenuForm = EventHandler.oApplication.Forms.ActiveForm;
                            if (oMenuForm.UniqueID == GlobalVariables.bldID)
                            {
                                GlobalVariables.oBilloflanding.MenuEvent(ref pVal, out BubbleEvent);
                                break;
                            }
                            //if (oMenuForm.UniqueID == GlobalVariables.bpID)
                            //{
                            //    GlobalVariables.oBPMapiing.MenuEvent(ref pVal, out BubbleEvent);
                            //    break;
                            //}
                            break;
                    }
                }

                if ((pVal.BeforeAction == false))
                {


                    if (pVal.MenuUID == GlobalVariables.bldID)
                    {
                        if (GlobalVariables.oGFun.FormExist(GlobalVariables.bldID))
                        {
                            oApplication.Forms.Item(GlobalVariables.bldID).Visible = true;
                            oApplication.Forms.Item(GlobalVariables.bldID).Select();
                        }
                        else
                        {
                            GlobalVariables.oBilloflanding.loadBillofLading();
                        }
                    }


                    GlobalVariables.oForm = oApplication.Forms.ActiveForm;
                    if ((pVal.MenuUID == "526"))
                    {
                        GlobalVariables.oCompany.Disconnect();
                        oApplication.StatusBar.SetText((GlobalVariables.addonName + " AddOn is DisConnected . . ."), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                }
            }
            catch (Exception ex)
            {
                oApplication.StatusBar.SetText(("Menu Event Failed : " + ex.Message), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

            }
            finally
            {
            }
        }
        #endregion

        #region Application Events
        private static void oApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            try
            {
                switch (EventType)
                {
                    case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                        System.Windows.Forms.Application.Exit();
                        break;
                }

            }
            catch (Exception ex)
            {
                oApplication.StatusBar.SetText("Application Event Failed: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            finally
            {
            }

        }
        #endregion

        #region " ... 2) Item Event ..."
        private static void oApplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {


                if (pVal.FormTypeEx == "143")
                {
                    GlobalVariables.ARinvoice.ItemEvent(pVal.FormUID, ref pVal, out BubbleEvent);
                }
                if (pVal.FormUID == GlobalVariables.BillWizardID)
                {
                    GlobalVariables.oBillWizard.ItemEvent(GlobalVariables.BillWizardID, ref pVal, out BubbleEvent);
                }
                if (pVal.FormUID == GlobalVariables.bldID)
                {
                    GlobalVariables.oBilloflanding.ItemEvent(GlobalVariables.bldID, ref pVal, out BubbleEvent);
                }

            }

            catch (Exception ex)
            {
                oApplication.StatusBar.SetText(("Item Event Failed : " + ex.Message), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            finally
            {
            }
        }
        #endregion

        #region " ... 3) FormDataEvent ..."
        private static void oApplication_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {

                if (BusinessObjectInfo.FormUID == GlobalVariables.bldID)
                {
                    GlobalVariables.oBilloflanding.FormDataEvent(ref BusinessObjectInfo, ref BubbleEvent);
                }
            }
            catch (Exception ex)
            {
                oApplication.StatusBar.SetText(("Form DataEvent Failed : " + ex.Message), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            finally
            {
            }
        }
        #endregion

        #region " ... 4) Status Bar Event ..."
        public static void oApplication_StatusBarEvent(string Text, SAPbouiCOM.BoStatusBarMessageType MessageType)
        {
            try
            {
                if (MessageType == SAPbouiCOM.BoStatusBarMessageType.smt_Warning | MessageType == SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                {
                    //System.Media.SystemSounds.Asterisk.Play();
                }
            }
            catch (Exception ex)
            {
                oApplication.StatusBar.SetText(GlobalVariables.addonName + " StatusBarEvent Event Failed : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            finally
            {
            }
        }
        #endregion

        #region " ... 5) Set Event Filter ..."
        public static void SetEventFilter()
        {
            try
            {
            }
            catch (Exception ex)
            {
                oApplication.StatusBar.SetText(ex.Message);
            }
            finally
            {
            }
        }
        #endregion

        #region " ... 6) Right Click Event ..."
        private static void oApplication_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
            }
            catch (Exception ex)
            {
                oApplication.StatusBar.SetText((GlobalVariables.addonName + (" : Right Click Event Failed : " + ex.Message)), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            finally
            {
            }
        }
        #endregion


    }
}
