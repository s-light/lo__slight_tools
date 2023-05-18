# -*- coding: utf-8 -*-
"""
Save the selection as sql text-file

based on
https://wiki.documentfoundation.org/Macros/Python_Guide/Documents
"""

import time
import threading
import os

import uno
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS

from com.sun.star.beans import PropertyValue

# import helper_functions as helper_fn

CTX = uno.getComponentContext()
SM = CTX.getServiceManager()


def create_instance(name, with_context=False):
    """Create Instance Helper Function."""
    if with_context:
        instance = SM.createInstanceWithContext(name, CTX)
    else:
        instance = SM.createInstance(name)
    return instance


def run_in_thread(func):
    """Run function in thread."""

    def run(*k, **kw):
        mythread = threading.Thread(target=func, args=k, kwargs=kw)
        mythread.start()
        return mythread

    return run


def msgbox(
    message,
    title="LibreOffice",
    buttons=MSG_BUTTONS.BUTTONS_OK,
    type_msg="infobox",
):
    """
    Create message box.

    type_msg: infobox, warningbox, errorbox, querybox, messbox

    https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XMessageBoxFactory.html
    """
    toolkit = create_instance("com.sun.star.awt.Toolkit")
    parent = toolkit.getDesktopWindow()
    mbox = toolkit.createMessageBox(
        parent,
        type_msg,
        buttons,
        title,
        str(message),
    )
    return mbox.execute()


def get_current_document():
    """Get current document."""
    desktop = create_instance("com.sun.star.frame.Desktop", True)
    doc = desktop.getCurrentComponent()
    # msgbox(doc.Title)
    # same as ??
    # doc = XSCRIPTCONTEXT.getDocument()
    return doc


def dict_to_property(values: dict, uno_any: bool = False):
    # source: https://wiki.documentfoundation.org/Macros/Python_Guide/Useful_functions#Dictionary_to_properties
    ps = tuple([PropertyValue(Name=n, Value=v) for n, v in values.items()])
    if uno_any:
        ps = uno.Any("[]com.sun.star.beans.PropertyValue", ps)
    return ps


@run_in_thread
def update_statusbar(statusbar, text, limit):
    """Update statusbar."""
    statusbar.start(text, limit)
    for i in range(limit):
        statusbar.setValue(i)
        time.sleep(1)
    # ~ Is important free status bar
    statusbar.end()


def statusbar_animate_progress(text, target=10):
    """Show Progress on statusbar."""
    doc = get_current_document()
    statusbar = doc.getCurrentController().getStatusIndicator()
    update_statusbar(statusbar, text, target)


######################################################


######################################################


def get_sheet_sql_file(doc):
    """append sheet name to filename and change ending to sql."""
    filename_full_path = uno.fileUrlToSystemPath(doc.URL)
    print("filename_full_path:")
    print(filename_full_path)
    print("", filename_full_path)

    # so extract the FileName from the filename_full_path
    path_full = os.path.dirname(filename_full_path)
    filename = os.path.basename(filename_full_path)
    file_basename = os.path.splitext(filename)[0]
    file_ext = os.path.splitext(filename)[1]
    # print("doc_name: " + doc_name)

    sheet_active = doc.CurrentController.ActiveSheet
    sheet_name = sheet_active.Name
    print("sheet_name: ", sheet_name)

    file_basename_new = file_basename + "__" + sheet_name
    # pdf_filename_full_path = path_full + "/" + file_basename_new + file_ext
    target_filename_full_path = os.path.join(path_full, file_basename_new) + ".sql"
    print("target_filename_full_path:", target_filename_full_path)
    return target_filename_full_path


def save_selection_as_sql(call_args=[]):
    """Save selection as sql text file."""
    print("-" * 42)
    print("save_selection_as_sql:")

    doc = get_current_document()
    # doc = XSCRIPTCONTEXT.getDocument()
    sheets = doc.Sheets
    sheet_active = doc.CurrentController.ActiveSheet
    selection = doc.CurrentController.Selection
    data = selection.DataArray


    sql_file = get_sheet_sql_file(doc)
    print("sql_file:", sql_file)
    # print("data:", data)
    try:
        with open(sql_file, 'w', encoding="utf-8") as f:
            for row in data:
                print("row:", row)
                f.write('\n        '.join(row))
                f.write('\n')
    except Exception as e:
        print("*" * 42)
        print("error", e)
        print("*" * 42)
    print("-" * 42)


#

g_exportedScripts = (save_selection_as_sql,)
