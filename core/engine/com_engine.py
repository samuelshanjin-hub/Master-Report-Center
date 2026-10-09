class ComEngine:
    """Optional Windows Excel COM adapter.

    This V2 Structure keeps COM isolated so production-grade formatting,
    hidden columns, validation, conditional formatting, shapes and pictures
    can be migrated without coupling those rules to the GUI.
    """
    @staticmethod
    def available():
        try:
            import win32com.client  # noqa: F401
            return True
        except Exception:
            return False

    @staticmethod
    def create_excel(visible=False):
        import win32com.client
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = visible
        excel.DisplayAlerts = False
        return excel
