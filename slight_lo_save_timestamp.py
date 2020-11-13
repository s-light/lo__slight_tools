# -*- coding: utf-8 -*-
"""Save the Active Document with new current timeStamp."""

# regex
import re

import time
import os.path

# libreoffice
import uno

# import helper_functions as helper_fn


############################################################################


# -*- coding: utf-8 -*-
"""
Save the Active Document with new current timeStamp.

based on
https://wiki.documentfoundation.org/Macros/Python_Guide/Documents
"""

import time
import threading

# libreoffice
import uno
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS

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
    return doc


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


######################################################


######################################################


def update_timestamp(file_basename, add_missing=True, new_post=False):
    """Update timestamp in filename."""
    # infos / tests:
    # get pre and post part of current name
    # Test_20150218_1852_sk
    # reTest = re.compile(r"^(.*_)2015")
    # res = reTest.search("Hallo_2015")
    # Gui.SendMsgToActiveView("Save")
    # App.ActiveDocument.Name
    # --> "Test_20150218_1852_sk"
    # App.getDocument("Test_20150218_1852_sk").save()
    # App.ActiveDocument.FileName
    # u'C:/_Local_DATA/_projects/2015_01_00__xxxx/Zeichnungen/Test_20150219_1106_sk.FCStd'
    re_timestamp = re.compile(r"^(.*_)20[0-9]{2}[0-9]{2}[0-9]{2}_[0-9]{4}(_.*$)?")
    re_timestamp_match = re_timestamp.search(file_basename)
    if re_timestamp_match:
        pre_part = re_timestamp_match.groups()[0]
        post_part = re_timestamp_match.groups()[1]
        if new_post:
            file_basename_new = pre_part + time.strftime("%Y%m%d_%H%M") + new_post
        else:
            file_basename_new = pre_part + time.strftime("%Y%m%d_%H%M") + post_part
    else:
        if add_missing:
            file_basename_new = file_basename + time.strftime("%Y%m%d_%H%M")
            if new_post:
                file_basename_new = file_basename_new + new_post
        else:
            file_basename_new = None
    return file_basename_new


def test(args):
    """Test."""
    doc = get_current_document()
    # filename_full_path = uno.fileUrlToSystemPath(doc.URL)
    # path_full = os.path.dirname(filename_full_path)
    # filename = os.path.basename(filename_full_path)
    # msgbox(filename)
    # user_profile = create_instance("/org.openoffice.UserProfile/Data", True)
    user_profile = create_instance("/org.openoffice.UserProfile", True)
    msgbox(user_profile)


def test_statusbar(args):
    """Test."""
    statusbar_animate_progress("Hello World", 20)


def save_active_doc_with_timestamp(args):
    """Save the Active Document with new current timeStamp."""
    print("-" * 42)
    print("save_active_doc_with_timestamp:")

    doc = get_current_document()
    filename_full_path = uno.fileUrlToSystemPath(doc.URL)
    print("filename_full_path:")
    print(filename_full_path)
    print("", filename_full_path)

    # so extract the FileName from the path_full
    path_full = os.path.dirname(filename_full_path)
    filename = os.path.basename(filename_full_path)
    file_basename = os.path.splitext(filename)[0]
    file_ext = os.path.splitext(filename)[1]
    # print("doc_name: " + doc_name)

    file_basename_new = update_timestamp(file_basename)
    if file_basename_new:
        # put together new filename_full_path_new
        filename_full_path_new = path_full + "/" + file_basename_new + file_ext
        print("filename_full_path_new:")
        print(filename_full_path_new)
        # https://wiki.openoffice.org/wiki/Documentation/DevGuide/OfficeDev/Storing_Documents
        doc.storeAsURL(uno.systemPathToFileUrl(filename_full_path_new), ())
    else:
        info = "no timestamp found in filename - so i can't update it."
        print(info)
        msgbox(info)

    print("-" * 42)


#

g_exportedScripts = (save_active_doc_with_timestamp, test)
