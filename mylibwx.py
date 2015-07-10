import wx
from wx.lib.dialogs import messageDialog, singleChoiceDialog


def open_file(wcd, message="Select a file", path=""):
    application = wx.PySimpleApp()
    # Create an open file dialog
    dialog = wx.FileDialog(None, message, defaultDir=path, wildcard=wcd, style=wx.OPEN)
    # Show the dialog and get user input
    if dialog.ShowModal() == wx.ID_OK:
        path = dialog.GetPath()
    # Destroy the dialog
    dialog.Destroy()
    return path


def saveas(wcd, message="Save file as...", path=""):
    application = wx.PySimpleApp()
    save_dlg = wx.FileDialog(None, message, defaultDir=path, wildcard=wcd, style=wx.SAVE | wx.OVERWRITE_PROMPT)
    if save_dlg.ShowModal() == wx.ID_OK:
        path = save_dlg.GetPath()

    save_dlg.Destroy()
    return path


def open_pyfile(message="Select a file", path=""):
    return open_file('Text File (*.txt)|*.txt|Python File (*.py)|*.py', message, path)


def open_zipfile(message="Select a file", path=""):
    return open_file('Zip file (*.zip)|*.zip', message, path)


def open_textfile(message="Select a .txt file", path=""):
    return open_file('Text File (*.txt)|*.txt', message, path)


def open_csvfile(message="Select a csv file", path=""):
    return open_file('Comma Separated Version (*.csv)|*.csv', message, path)


def open_xls(message="Select a spreadsheet file", path=""):
    return open_file('Excel 97/2003 Spreadsheet (*.xls)|*.xls', message, path)


def open_xlspreadsheet(message="Select a spreadsheet file", path=""):
    wcd = 'Excel 2007 Spreadsheet (*.xlsx)|*.xlsx|' \
          'Excel 97/2003 Spreadsheet (*.xls)|*.xls|' \
          'Comma Separated Version (*.csv)|*.csv'
    return open_file(wcd=wcd, message=message, path=path)


def open_xlsx(message="Select a spreadsheet file", path=""):
    wcd = 'Excel 2007 Spreadsheet (*.xlsx)|*.xlsx'
    return open_file(wcd=wcd, message=message, path=path)


def open_xlsm(message="Select a spreadsheet file", path=""):
    wcd = 'Excel 2007 Spreadsheet (*.xlsm)|*.xlsm'
    return open_file(wcd=wcd, message=message, path=path)


def saveas_xlsx(path="", mesg='Save file as...'):
    wcd = 'Excel 2007 Spreadsheet (*.xlsx)|*.xlsx'
    return saveas(wcd=wcd, message=mesg, path=path)


def saveas_textfile(path=""):
    return saveas('Text File (*.txt)|*.txt', path)


def saveas_xls(path=""):
    return saveas('Excel 97/2003 Spreadsheet (*.xls)|*.xls', path)


def saveas_csv(path=""):
    return saveas('Comma Separated Version (*.csv)|*.csv', path)


def selectDir(message="Select a folder", path="c:/"):
    path = ""
    application = wx.PySimpleApp()
    # Create an open file dialog
    dialog = wx.DirDialog(None, message, style=1, defaultPath=path, pos=(10, 10))
    # Show the dialog and get user input
    if dialog.ShowModal() == wx.ID_OK:
        path = dialog.GetPath()
    # Destroy the dialog
    dialog.Destroy()
    return path


def OkDialog(message, caption):
    application = wx.PySimpleApp()
    messageDialog(None, message, caption, wx.OK)


def ErrorDialog(message, caption):
    application = wx.PySimpleApp()
    messageDialog(None, message, caption, wx.OK | wx.ICON_ERROR)


def YesNoDialog(message, caption):
    application = wx.PySimpleApp()
    dlg = messageDialog(None, message, caption, wx.YES | wx.NO | wx.ICON_INFORMATION)
    return dlg.returnedString


def YesNoCancelDialog(message, caption):
    application = wx.PySimpleApp()
    dlg = messageDialog(None, message, caption, wx.YES | wx.NO | wx.CANCEL | wx.ICON_INFORMATION)
    return dlg.returnedString


def ChoiceDialog(message, caption, choices=None):
    """
    display a dialog box with drop down list containing items in a list,
    return the item selected by user, return None if click Cancel

    reconcile expectation of singleChoiceDialog by converting choices into a list
    of strings if choices is a list containing something else.
    """
    if choices is None:
        choices = []
    t_choices = dict((str(i), i) for i in choices)
    user_choice = ''
    selection = None
    app = wx.PySimpleApp()
    dialog = singleChoiceDialog(None, message, caption, map(str, choices))
    # using map(str, choices) instead of t_choices.keys() because former list is sorted
    if dialog.accepted:
        user_choice = str(dialog.selection)
        # pick up the corresponding item from the list
        # this may be necessary if the choices do not contain strings
        selection = t_choices[user_choice]
    # The user exited the dialog without pressing the "OK" button
    else:
        pass
    return selection


def MChoiceDialog(message, caption, choices=None):
    """
    display a dialog box with drop down list containing items in a list,
    return the item selected by user, return None if click Cancel

    reconcile expectation of singleChoiceDialog by converting choices into a list
    of strings if choices is a list containing something else.
    """
    if choices is None:
        choices = []
    t_choices = dict((str(i), i) for i in choices)
    user_choice = ''
    selection = None
    app = wx.PySimpleApp()
    dialog = wx.MultiChoiceDialog(None, message, caption, map(str, choices))
    # dialog = singleChoiceDialog ( None, message, caption, map(str, choices))
    # using map(str, choices) instead of t_choices.keys() because former list is sorted
    if dialog.ShowModal() == wx.ID_OK:
        selection = [choices[r] for r in dialog.GetSelections()]
        # pick up the corresponding item from the list
        # this may be necessary if the choices do not contain strings
    # The user exited the dialog without pressing the "OK" button
    else:
        pass
    return selection


def TextDialog(question, caption):
    message = None
    app = wx.PySimpleApp()
    dlg = wx.TextEntryDialog(None, question, caption, '')
    if dlg.ShowModal() == wx.ID_OK:
        message = dlg.GetValue()    # this line should be indented
        dlg.Destroy()
    else:
        pass
    return message


if __name__=='__main__':
    print YesNoDialog("test", "yesno")
