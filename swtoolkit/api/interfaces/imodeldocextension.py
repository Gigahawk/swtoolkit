import win32com
import pythoncom

from ..custompropertymanager import CustomPropertyManager


class IModelDocExtension:
    def __init__(self, parent):
        self._instance = parent.Extension

    def custom_property_manager(self, config_name):
        return CustomPropertyManager(self._instance, config_name)

    def rebuild(self, options):
        arg = win32com.client.VARIANT(pythoncom.VT_I4, options)
        return self._instance.Rebuild(arg)

    def select_by_id2(self):
        pass

    def view_zoom_to_sheet(self):
        pass

    def save_as2(self, name, export_file_data, options):
        arg1 = win32com.client.VARIANT(pythoncom.VT_BSTR, name)
        arg2 = win32com.client.VARIANT(pythoncom.VT_I4, 0) # swSaveAsVersion_e
        arg3 = win32com.client.VARIANT(pythoncom.VT_I4, options) # swSaveAsOptions_e
        arg4 = win32com.client.VARIANT(pythoncom.VT_VARIANT, export_file_data) # IExportPdfData
        arg5 = win32com.client.VARIANT(pythoncom.VT_BSTR, "") # ReferencePrefixOrSuffixText
        arg6 = win32com.client.VARIANT(pythoncom.VT_BOOL, False) # AddTextAsPrefix
        arg7 = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4, None # Errors
        )
        arg8 = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4, None # Warnings
        )
        retval = self._instance.SaveAs2(
            arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8
        )
        return retval, arg7, arg8

    def get_advanced_save_as_options(self, options=0):
        arg = win32com.client.VARIANT(pythoncom.VT_I4, options)
        return self._instance.GetAdvancedSaveAsOptions(arg)

    def save_as3(self, name, export_file_data, options):
        advanced_save_as_options = self.get_advanced_save_as_options()
        arg1 = win32com.client.VARIANT(pythoncom.VT_BSTR, name)
        arg2 = win32com.client.VARIANT(pythoncom.VT_I4, 0) # swSaveAsVersion_e
        arg3 = win32com.client.VARIANT(pythoncom.VT_I4, options) # swSaveAsOptions_e
        arg4 = win32com.client.VARIANT(pythoncom.VT_VARIANT, export_file_data) # IExportPdfData
        # TODO: Figure out why this causes a type mismatch error
        arg5 = win32com.client.VARIANT(pythoncom.VT_VARIANT, advanced_save_as_options) # IAdvancedSaveAsOptions
        arg6 = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4, None # Errors
        )
        arg7 = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4, None # Warnings
        )
        retval = self._instance.SaveAs3(
            arg1, arg2, arg3, arg4, arg5, arg6, arg7
        )
        return retval, arg6, arg7

    def save_pack_and_go(self):
        pass

    def set_user_preference_toggle(self):
        pass
