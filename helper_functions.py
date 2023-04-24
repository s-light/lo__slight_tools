# -*- coding: utf-8 -*-
"""
Helper Functions..

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
def statusbar_progress_update(statusbar, text, limit):
    """Update statusbar."""
    limit = limit * 100
    statusbar.start(text, limit)
    for i in range(limit):
        statusbar.setValue(i)
        time.sleep(0.01)
    # ~ Is important free status bar
    statusbar.end()


def statusbar_animate_progress(text, target=10):
    """Show Progress on statusbar."""
    doc = get_current_document()
    statusbar = doc.getCurrentController().getStatusIndicator()
    statusbar_progress_update(statusbar, text, target)
