"""
PyXLL Examples: Menus

The PyXLL Excel Addin is configured to load one or more
python modules when it's loaded.

Menus can be added to Excel via the pyxll xl_menu decorator.
"""

import logging
_log = logging.getLogger(__name__)

# the webbrowser module is used in an example to open the log file
try:
    import webbrowser
except ImportError:
    _log.warning("*** webbrowser could not be imported             ***")
    _log.warning("*** the menu examples will not work correctly    ***")

import os

#
# 1) Basics - adding a menu items to Excel
#

#
# xl_menu is the decorator used for addin menus to Excel.
#
from pyxll import xl_menu, get_config, xlcAlert

#
# The only required argument is the menu item name.
# The example below will add a new menu item to the
# addin's default menu.
#

@xl_menu("Example Menu Item 1")
def on_example_menu_item_1():
    xlcAlert("Hello from PyXLL")

#
# menu items are normally sorted alphabetically, but the order
# keyword can be used to influence the ordering of the items
# in a menu.
#
# The default value for all sort keyword arguments is 0, so positive
# values will result in the item appearing further down the list
# and negative numbers result in the item appearing further up.
#

@xl_menu("Another example menu item", order=1)
def on_example_menu_item_2():
    xlcAlert("Hello again from PyXLL")

#
# It's possible to add items to menus other than the default menu.
# The example below creates a new menu called 'My new menu' with
# one item 'Click me' in it.
#
# The menu_order keyword is optional, but may be used to influence
# the order that the custom menus appear in.
#

@xl_menu("Click me", menu="PyXLL example menu", menu_order=1)
def on_example_menu_item_3():
    xlcAlert("Wow, a different menu!")

#
# 2) Sub-menus
#

# it's possible to add sub-menus just by using the sub_menu
# keyword argument. The example below adds a new sub menu
# 'Sub Menu' to the default menu.
#
# The order keyword argument affects where the sub menu will
# appear in the parent menu, and the sub_order keyword argument
# affects where the item will appear in the sub menu.
#

@xl_menu("Click me", sub_menu="More Examples", order=2)
def on_example_submenu_item_1():
    xlcAlert("Sub-menus can be created easily with PyXLL")

#
# A simple menu item to show how to get the PyXLL config
# object and open the log file.
#

@xl_menu("Open log file", order=3)
def on_open_logfile():
    # the PyXLL config is accessed as a ConfigParser.ConfigParser object
    config = get_config()
    if config.has_option("LOG", "path") and config.has_option("LOG", "file"):
        path = os.path.join(config.get("LOG", "path"), config.get("LOG", "file"))
        webbrowser.open("file://%s" % path)