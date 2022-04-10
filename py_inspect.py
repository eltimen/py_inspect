from PyQt5.QtCore import QLocale
from PyQt5.QtCore import QCoreApplication
from PyQt5.QtCore import QSettings
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QAbstractTableModel
from PyQt5.QtCore import QVariant
from PyQt5.QtGui import QStandardItemModel
from PyQt5.QtGui import QStandardItem
from PyQt5.QtGui import QIntValidator
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QGridLayout
from PyQt5.QtWidgets import QHBoxLayout
from PyQt5.QtWidgets import QVBoxLayout
from PyQt5.QtWidgets import QLabel
from PyQt5.QtWidgets import QComboBox
from PyQt5.QtWidgets import QTreeView
from PyQt5.QtWidgets import QTableView
from PyQt5.QtWidgets import QMenuBar
from PyQt5.QtWidgets import QAction
from PyQt5.QtWidgets import QListView
from PyQt5.QtWidgets import QDialog
from PyQt5.QtWidgets import QLineEdit
from PyQt5.QtWidgets import QGroupBox
from PyQt5.QtWidgets import QRadioButton
from PyQt5.QtWidgets import QVBoxLayout
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QCheckBox
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import QTextEdit
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QHeaderView
import sys
import warnings

warnings.simplefilter("ignore", UserWarning)
sys.coinit_flags = 2
# TODO fix imports
import inspect
import pywinauto
from pywinauto import backend
import win32api
import threading
import time

# TODO rewrite with hooks + find in tree
def lookForMouse(w):
    desktop = pywinauto.Desktop(backend='uia')
    while True:
        time.sleep(0.5)
        if w.mouse.isChecked():
            x, y = win32api.GetCursorPos()
            desktop.from_point(x, y).draw_outline()

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    w = MyWindow()
    w.show()
    #x = threading.Thread(target=lookForMouse, args=(w,), daemon=True)
    # x.start()
    sys.exit(app.exec_())


class InfoDialog(QDialog):
    def __init__(self, title, items, parent=None):
        super().__init__(parent)

        self.setWindowTitle(title)
        listView = QListView()
        model = QStandardItemModel()
        listView.setModel(model)

        for item in items:
            model.appendRow(QStandardItem(str(item)))

        self.layout = QGridLayout()
        self.layout.addWidget(listView)
        self.setLayout(self.layout)
        self.resize(400, 300)


class ClickInput(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.button = QGroupBox("button")
        self.left = QRadioButton("left")
        self.right = QRadioButton("right")
        self.middle = QRadioButton("middle")
        self.coords = QGroupBox("coords")
        self.x = QLineEdit()
        self.x.setValidator(QIntValidator())
        self.y = QLineEdit()
        self.y.setValidator(QIntValidator())
        self.double = QCheckBox('double')
        self.wheel = QGroupBox('wheel_dist')
        self.wheel_dist = QLineEdit()
        self.wheel_dist.setValidator(QIntValidator())
        self.confirm = QPushButton('confirm')
        self.setup_ui()

    def setup_ui(self):

        self.setWindowTitle('click_input()')

        grid = QGridLayout()

        self.button.setFont(QFont("Sanserif", 14))

        vbox1 = QVBoxLayout()
        vbox1.addWidget(self.left)
        vbox1.addWidget(self.right)
        vbox1.addWidget(self.middle)

        vbox1.addStretch(1)
        self.button.setLayout(vbox1)

        self.coords.setFont(QFont("Sanserif", 14))

        vbox2 = QVBoxLayout()
        vbox2.addWidget(self.x)
        vbox2.addWidget(self.y)

        vbox2.addStretch(1)
        self.coords.setLayout(vbox2)

        self.double.setChecked(False)

        vbox3 = QVBoxLayout()
        vbox3.addWidget(self.wheel_dist)

        # vbox3.addStretch(1)
        self.wheel.setLayout(vbox3)

        grid.addWidget(self.button, 0, 0)
        grid.addWidget(self.coords, 0, 1)
        grid.addWidget(self.double, 1, 0)
        grid.addWidget(self.wheel, 1, 1)
        grid.addWidget(self.confirm, 2, 0, 1, 2)

        self.setLayout(grid)

        self.resize(400, 300)

        self.confirm.clicked.connect(self.win_close)

    def win_close(self):
        if self.x.text() == '' and self.y.text() != '' or self.x.text() != '' and self.y.text() == '':
            QMessageBox.warning(self, "Attention",
                                "Please type coords correctly!")
        else:
            self.close()


class MyWindow(QWidget):
    def __init__(self):
        super(MyWindow, self).__init__()

        # Methods
        # print(set([attr for attr in dir(pywinauto.base_wrapper.BaseWrapper)if not attr.startswith('_')]))
        self.base_methods = {
            # TODO remove overriden methods
            #'capture_as_image': self.__capture_as_image,
            'children': self.__children,
            'click_input': self.__click_input,
            'close': self.__close,
            'descendants': self.__descendants,
            #'draw_outline': self.__draw_outline,
            'set_focus': self.__set_focus,
            'texts': self.__texts,
            'type_keys': self.__type_keys,
            'wait_enabled': self.__wait_enabled,
            'wait_not_enabled': self.__wait_not_enabled,
            'wait_not_visible': self.__wait_not_visible,
            'wait_visible': self.__wait_visible,
            'window_text': self.__window_text
        }
        self.backend_methods = {
            'win32': {
                # element_info_win = pywinauto.backend.registry.backends['win32'].element_info_class()
                # print(set([attr for attr in dir(pywinauto.controls.hwndwrapper.HwndWrapper) if not attr.startswith('_') and not attr[0].isupper() and inspect.ismethod(getattr(pywinauto.controls.hwndwrapper.HwndWrapper(element_info_win), attr))]))
                'backend': {
                    'full_control_type': None,
                    'set_keyboard_focus': None,
                    'style': None,
                    'set_transparency': None,
                    'context_help_id': None,
                    'popup_window': None,
                    'has_exstyle': None,
                    'control_type': None,
                    'maximize': None,
                    'is_minimized': None,
                    'has_focus': None,
                    'fonts': None,
                    'owner': None,
                    'double_click': None,
                    'font': None,
                    'send_command': None,
                    'menu_item': None,
                    'send_message': None,
                    'menu_select': None,
                    'handle': None,
                    'client_rect': None,
                    'set_application_data': None,
                    'click': None,
                    'right_click': None,
                    'client_rects': None,
                    'send_keystrokes': None,
                    'debug_message': None,
                    'press_mouse': None,
                    'has_keyboard_focus': None,
                    'post_message': None,
                    'move_window': None,
                    'minimize': None,
                    'is_unicode': None,
                    'set_window_text': None,
                    'post_command': None,
                    'exstyle': None,
                    'scroll': None,
                    'get_toolbar': None,
                    'get_focus': None,
                    'is_normal': None,
                    'menu': None,
                    'restore': None,
                    'send_chars': None,
                    'is_maximized': None,
                    'is_active': None,
                    'drag_mouse': None,
                    'get_show_state': None,
                    'has_style': None,
                    'user_data': None,
                    'release_mouse': None,
                    'move_mouse': None,
                    'notify_parent': None,
                    'send_message_timeout': None,
                    'menu_items': None
                },
                # print(pywinauto.controls.hwndwrapper.HwndMeta.str_wrappers)
                'controls': {
                    'hwndwrapper.DialogWrapper': {'client_area_rect': None, 'force_close': None, 'show_in_taskbar': None, 'hide_from_taskbar': None, 'run_tests': None, 'write_to_xml': None, 'is_in_taskbar': None},
                    'win32_controls.ButtonWrapper': {'uncheck_by_click_input': None, 'uncheck_by_click': None, 'check_by_click_input': None, 'get_check_state': None, 'is_checked': None, 'uncheck': None, 'check': None, 'check_by_click': None, 'set_check_indeterminate': None},
                    'win32_controls.ComboBoxWrapper': {'selected_index': None, 'select': None, 'item_count': None, 'dropped_rect': None, 'item_data': None, 'selected_text': None, 'item_texts': None},
                    'win32_controls.ListBoxWrapper': {'set_item_focus': None, 'select': None, 'item_count': None, 'selected_indices': None, 'item_data': None, 'item_rect': None, 'item_texts': None, 'is_single_selection': None, 'get_item_focus': None},
                    'win32_controls.EditWrapper': {'get_line': None, 'select': None, 'line_length': None, 'selection_indices': None, 'line_count': None, 'set_text': None, 'text_block': None, 'set_edit_text': None},
                    'win32_controls.StaticWrapper': {},
                    'win32_controls.PopupMenuWrapper': {},
                    'common_controls.ListViewWrapper': {'items': None, 'get_header_control': None, 'columns': None, 'select': None, 'item_count': None, 'is_focused': None, 'get_column': None, 'get_item_rect': None, 'is_checked': None, 'uncheck': None, 'check': None, 'is_selected': None, 'get_selected_count': None, 'column_widths': None, 'column_count': None, 'item': None, 'get_item': None, 'deselect': None},
                    'common_controls.TreeViewWrapper': {'ensure_visible': None, 'select': None, 'print_items': None, 'item_count': None, 'is_selected': None, 'item': None, 'tree_root': None, 'get_item': None, 'roots': None},
                    'common_controls.HeaderWrapper': {'get_column_text': None, 'get_column_rectangle': None, 'item_count': None},
                    'common_controls.StatusBarWrapper': {'border_widths': None, 'get_part_text': None, 'get_part_rect': None, 'part_right_edges': None, 'part_count': None},
                    'common_controls.TabControlWrapper': {'row_count': None, 'get_tab_text': None, 'tab_count': None, 'get_tab_rect': None, 'select': None, 'get_selected_tab': None},
                    'common_controls.ToolbarWrapper': {'get_button': None, 'get_button_struct': None, 'check_button': None, 'press_button': None, 'get_tool_tips_control': None, 'tip_texts': None, 'get_button_rect': None, 'button': None, 'menu_bar_click_input': None, 'button_count': None},
                    'common_controls.ReBarWrapper': {'get_tool_tips_control': None, 'get_band': None, 'band_count': None},
                    'common_controls.ToolTipsWrapper': {'tool_count': None, 'get_tip': None, 'get_tip_text': None},
                    'common_controls.UpDownWrapper': {'decrement': None, 'set_base': None, 'increment': None, 'set_value': None, 'get_value': None, 'get_base': None, 'get_buddy_control': None, 'get_range': None},
                    'common_controls.TrackbarWrapper': {'set_range_min': None, 'get_line_size': None, 'set_position': None, 'set_line_size': None, 'get_sel_end': None, 'get_page_size': None, 'get_position': None, 'get_sel_start': None, 'set_page_size': None, 'get_num_ticks': None, 'get_tooltips_control': None, 'get_channel_rect': None, 'set_sel': None, 'get_range_min': None, 'get_range_max': None, 'set_range_max': None},
                    'common_controls.AnimationWrapper': {},
                    'common_controls.ComboBoxExWrapper': {},
                    'common_controls.DateTimePickerWrapper': {'set_time': None, 'get_time': None},
                    'common_controls.HotkeyWrapper': {},
                    'common_controls.IPAddressWrapper': {},
                    'common_controls.CalendarWrapper': {'get_view': None, 'set_today': None, 'get_id': None, 'set_current_date': None, 'set_border': None, 'get_today': None, 'get_current_date': None, 'set_day_states': None, 'hit_test': None, 'set_first_weekday': None, 'set_id': None, 'get_month_range': None, 'set_color': None, 'set_view': None, 'calc_min_rectangle': None, 'get_border': None, 'set_month_delta': None, 'place_in_calendar': None, 'get_month_delta': None, 'count': None, 'get_first_weekday': None},
                    'common_controls.PagerWrapper': {'set_position': None, 'get_position': None},
                    'common_controls.ProgressWrapper': {'set_position': None, 'get_step': None, 'step_it': None, 'get_state': None, 'get_position': None}
                }
            },
            'uia': {
                # element_info_uia = pywinauto.backend.registry.backends['uia'].element_info_class()
                # print(set([attr for attr in dir(pywinauto.controls.uiawrapper.UIAWrapper) if not attr.startswith('_') and not attr.startswith('iface_') and inspect.ismethod(getattr(pywinauto.controls.uiawrapper.UIAWrapper(element_info_uia), attr))]))
                'backend': {
                    'is_selection_required': None,
                    'expand': None,
                    'maximize': None,
                    'is_minimized': None,
                    'menu_select': None,
                    'is_selected': None,
                    'legacy_properties': None,
                    'collapse': None,
                    'set_value': None,
                    'is_keyboard_focusable': None,
                    'has_keyboard_focus': None,
                    'is_expanded': None,
                    'move_window': None,
                    'minimize': None,
                    'selected_item_index': None,
                    'scroll': None,
                    'invoke': None,
                    'can_select_multiple': None,
                    'is_normal': None,
                    'restore': None,
                    'children_texts': None,
                    'is_maximized': None,
                    'is_active': None,
                    'is_collapsed': None,
                    'get_selection': None,
                    'get_show_state': None,
                    'get_expand_state': None,
                    'select': None
                },
                # print(pywinauto.controls.uiawrapper.UiaMeta.control_type_to_cls)
                # pywinauto.windows.uia_defines._build_pattern_ids_dic to check supported controls
                'controls': {
                    'uia_controls.WindowWrapper': {},
                    'uia_controls.ButtonWrapper': {'get_toggle_state': None, 'toggle': None, 'click': None},
                    'uia_controls.ComboBoxWrapper': {'item_count': None, 'selected_index': None, 'is_editable': None, 'selected_text': None},
                    'uia_controls.EditWrapper': {'selection_indices': None, 'set_edit_text': None, 'set_window_text': None, 'is_editable': None, 'get_value': None, 'line_count': None, 'set_text': None, 'text_block': None, 'line_length': None, 'get_line': None},
                    'uia_controls.TabControlWrapper': {'tab_count': None, 'get_selected_tab': None},
                    'uia_controls.SliderWrapper': {'small_change': None, 'value': None, 'min_value': None, 'max_value': None, 'large_change': None},
                    'uia_controls.HeaderWrapper': {},
                    'uia_controls.HeaderItemWrapper': {},
                    'uia_controls.ListItemWrapper': {'is_checked': None},
                    'uia_controls.ListViewWrapper': {'item_count': None, 'columns': None, 'items': None, 'get_column': None, 'get_header_controls': None, 'get_items': None, 'get_header_control': None, 'cell': None, 'item': None, 'get_item': None, 'column_count': None, 'cells': None, 'get_item_rect': None, 'get_selected_count': None},
                    'uia_controls.MenuItemWrapper': {'items': None},
                    'uia_controls.MenuWrapper': {'item_by_path': None, 'items': None, 'item_by_index': None},
                    'uia_controls.TooltipWrapper': {},
                    'uia_controls.ToolbarWrapper': {'button': None, 'button_count': None, 'buttons': None, 'check_button': None},
                    'uia_controls.TreeItemWrapper': {'get_child': None, 'is_checked': None, 'ensure_visible': None, 'sub_elements': None},
                    'uia_controls.TreeViewWrapper': {'item_count': None, 'get_item': None, 'roots': None, 'print_items': None},
                    'uia_controls.StaticWrapper': {}
                }
            },
            # HOW TO ADD BACKEND
            'other backend': {
                'backend': {
                    'other backend method 1': 'self.implementing_function_name',
                    'other backend method 2': 'self.implementing_function_name',
                },
                'controls': {
                    'other backend control wrapper 1': {'unique method 1': 'self.implementing_function_name', 'method 2': 'self.implementing_function_name'},
                    'other backend control wrapper 2': {'unique method 1': 'self.implementing_function_name', 'method 2': 'self.implementing_function_name'}
                }
            }
        }
        
        # TODO try resize instead
        self.setMinimumSize(1024, 768)
        self.setLocale(QLocale(QLocale.English, QLocale.UnitedStates))
        self.setWindowTitle(
            QCoreApplication.translate("MainWindow", "PyInspect"))

        self.settings = QSettings('py_inspect', 'MainWindow')

        self.current_elem_wrapper = None

        # Main layout
        self.mainLayout = QVBoxLayout()

        # Menu bar
        self.menu_bar = QMenuBar(self)
        self.action = self.menu_bar.addMenu("Actions")
        refresh = QAction('Refresh', self)
        refresh.triggered.connect(self.__refresh)
        self.mouse = QAction('Find by mouse', self)
        self.mouse.setCheckable(True)
        default = QAction('Default Action', self)
        default.triggered.connect(self.__default)
        self.action.addAction(refresh)
        self.action.addAction(self.mouse)
        self.action.addAction(default)
        self.action.addSeparator()

        self.bmethods = self.action.addMenu("Base Wrapper Methods")

        self.hmethods = self.action.addMenu("Hwnd Wrapper Methods")
        self.hmethods.menuAction().setVisible(False)
        self.umethods = self.action.addMenu("UIA Wrapper Methods")
        self.umethods.menuAction().setVisible(False)
        self.backend_menus = {
            'last_used': self.umethods,
            'win32': self.hmethods,
            'uia': self.umethods,
            # HOW TO ADD BACKEND
            'other backend': 'self.other_backend_menu'
        }

        self.cmethods = self.action.addMenu('Current Wrapper Unique Methods')
        self.backend_wrappers = {
            'win32': 'hwndwrapper.HwndWrapper',
            'uia': 'uiawrapper.UIAWrapper',
            # HOW TO ADD BACKEND
            'other backend': 'other backend wrapper name'
        }
        self.backend_inits = {
            'win32': pywinauto.controls.hwndwrapper.HwndWrapper,
            'uia': pywinauto.controls.uiawrapper.UIAWrapper,
            # HOW TO ADD BACKEND
            'other backend': 'other backend wrapper class name'
        }

        # Backend label
        self.backendLabel = QLabel("Backend Type")

        # Backend combobox
        self.comboBox = QComboBox()
        self.comboBox.setMouseTracking(False)
        self.comboBox.setMaxVisibleItems(5)
        self.comboBox.setObjectName("comboBox")
        self.comboBox.activated[str].connect(self.__show_tree)
        for _backend in backend.registry.backends.keys():
            self.comboBox.addItem(_backend)
        self.comboBox.setCurrentText('uia')

        # Tree view
        self.tree_view = QTreeView()
        self.tree_view.setColumnWidth(0, 150)
        self.tree_view.clicked.connect(self.__show_property)
        self.__initialize_calc()

        # Code generator
        self.editLabel = QLabel('Code generator')
        self.clear = QPushButton('Clear')
        self.clear.clicked.connect(self.__clear_edit)
        self.save = QPushButton('Save to file')
        self.save.clicked.connect(self.__save_edit)
        self.edit = QTextEdit()
        self.used_apps = {}

        # Table view
        self.table_view = QTableView()

        # Add top widgets to main window
        self.grid_tree = QGridLayout()
        self.grid_tree.addWidget(self.backendLabel, 0, 0, 1, 1)
        self.grid_tree.addWidget(self.comboBox, 0, 1, 1, 1)
        self.grid_tree.addWidget(self.tree_view, 1, 0, 1, 2)
        self.tree = QGroupBox('Controls View')
        self.tree.setLayout(self.grid_tree)

        self.grid_table = QGridLayout()
        self.grid_table.addWidget(self.editLabel, 0, 0, 1, 1)
        self.grid_table.addWidget(self.clear, 0, 1, 1, 1)
        self.grid_table.addWidget(self.save, 0, 2, 1, 1)
        self.grid_table.addWidget(self.edit, 1, 0, 1, 3)
        self.grid_table.addWidget(self.table_view, 2, 0, 1, 3)
        self.table = QGroupBox('Properties View')
        self.table.setLayout(self.grid_table)

        self.mainLayout.addWidget(self.menu_bar)
        self.hbox = QHBoxLayout()
        self.hbox.addWidget(self.tree)
        self.hbox.addWidget(self.table)
        self.mainLayout.addLayout(self.hbox)
        self.setLayout(self.mainLayout)
        geometry = self.settings.value('Geometry', bytes('', 'utf-8'))
        self.restoreGeometry(geometry)

    def __initialize_calc(self, _backend='uia'):
        self.element_info \
            = backend.registry.backends[_backend].element_info_class()
        self.tree_model = MyTreeModel(self.element_info, _backend)
        self.tree_model.setHeaderData(0, Qt.Horizontal, 'Controls')
        self.tree_view.setModel(self.tree_model)

    def __show_tree(self, text):
        backend = text
        self.current_elem_wrapper = None
        self.__initialize_calc(backend)

    def __show_property(self, index=None):
        data = index.data()
        current_backend = self.comboBox.currentText()
        self.current_elem_info = self.tree_model.info_dict.get(data)
        self.current_elem_wrapper = self.backend_inits[current_backend](self.current_elem_info)

        self.bmethods.clear()
        for method in self.base_methods.keys():
            action = QAction(method + '()', self)
            action.triggered.connect(self.base_methods[method])
            self.bmethods.addAction(action)

        self.backend_menus['last_used'].menuAction().setVisible(False)
        self.backend_menus['last_used'] = self.backend_menus[current_backend]
        self.backend_menus[current_backend].clear()
        self.backend_menus[current_backend].menuAction().setVisible(True)
        for method in self.backend_methods[current_backend]['backend'].keys():
                # if while not all implemented
                if self.backend_methods[current_backend]['backend'][method] != None:
                    action = QAction(method + '()', self)
                    action.triggered.connect(self.backend_methods[current_backend]['backend'][method])
                    self.backend_menus[current_backend].addAction(action)

        wrapper = str(self.current_elem_wrapper).split('-')[0][:-1]
        if wrapper != self.backend_wrappers[current_backend]:
            self.cmethods.clear()
            if wrapper in self.backend_methods[current_backend]['controls'].keys():
                for method in self.backend_methods[current_backend]['controls'][wrapper].keys():
                    # if while not all implemented
                    if self.backend_methods[current_backend]['controls'][wrapper][method] != None:
                        action = QAction(method + '()', self)
                        action.triggered.connect(self.backend_methods[current_backend]['controls'][wrapper][method])
                        self.cmethods.addAction(action)
            else:
                print('Unknown wrapper: ' + wrapper)

        # Debug
        #print(self.current_elem_info)
        #print(self.current_elem_wrapper)

        self.table_model \
            = MyTableModel(self.tree_model.props_dict.get(data), self)
        self.table_view.wordWrap()
        self.table_view.setModel(self.table_model)
        header = self.table_view.horizontalHeader()       
        header.setSectionResizeMode(1, QHeaderView.Stretch)
    
    def __clear_edit(self):
        self.edit.clear()
        self.used_apps = {}

    def __save_edit(self):
        save_file = QFileDialog.getSaveFileName(None, 'SaveTextFile','/', "Python Script (*.py)")
        text = self.edit.toPlainText()
        if save_file[0]: 
            with open(save_file[0], 'w') as file:
                file.write(text)
    
    def __write_method(self, method, ret):
        if len(self.used_apps) == 0:
            self.edit.append('from pywinauto.application import Application\n')
        top_parent = self.current_elem_info.top_level_parent
        elem_info = self.current_elem_info
        path = [elem_info]
        while elem_info.parent != top_parent:
            path.append(elem_info.parent)
            elem_info = elem_info.parent

        # check if already used
        if top_parent.process_id not in self.used_apps.keys():
            self.used_apps[top_parent.process_id] = 'app_{}'.format(len(self.used_apps) + 1)
            # TODO start correct app name
            self.edit.append(self.used_apps[top_parent.process_id] + ' = Application(backend="{}").start("{}")\n'.format(self.comboBox.currentText(), 'type correct argument for start()'))
        command = str(self.used_apps[top_parent.process_id])
        for ctrl in path[::-1]:
            if ctrl.name != '':
                ctrl_name = ctrl.name
            elif ctrl.control_type != '':
                ctrl_name = ctrl.control_type
            elif ctrl.class_name != '':
                ctrl_name = ctrl.class_name
            # TODO remove
            else: print('houston, we have a problem')
            command+= '["' + ctrl_name + '"]'
        command+= '.' + method
        if ret:
            self.edit.append('print(' + command + ')\n')
        else:
            self.edit.append(command + '\n')
    
    def closeEvent(self, event):
        geometry = self.saveGeometry()
        self.settings.setValue('Geometry', geometry)
        super(MyWindow, self).closeEvent(event)

    # Actions
    def __refresh(self):
        self.current_elem_wrapper = None
        self.__initialize_calc(str(self.comboBox.currentText()))

    def __default(self):
        # TODO add write method?
        if self.comboBox.currentText() == 'uia':
            if self.current_elem_info.legacy_action != '':
                self.current_elem_wrapper.iface_legacy_iaccessible.DoDefaultAction()
            else:
                self.current_elem_wrapper.set_focus()
        else:
            # TODO find default action in win32
            pass
        # HOW TO ADD BACKEND
        # just add default action implementation

    # Base Wrapper Methods
    #def __capture_as_image(self):
    #    img = self.current_elem_wrapper.capture_as_image()
    #    if img != None:
    #        img.show()
    #    else:
    #        print('can not capture as image')

    def __children(self):
        dlg = InfoDialog(
            'children()', self.current_elem_wrapper.children(), self)
        dlg.exec()
        self.__write_method('children()', True)

    # TODO there are more arguments in actual click_input
    def __click_input(self):
        dlg = ClickInput(self)
        dlg.exec()
        button = 'left'
        coords = (None, None)
        double = False
        wheel_dist = 0
        if dlg.right.isChecked():
            button = 'right'
        elif dlg.middle.isChecked():
            button = 'middle'
        if dlg.x.text() != '':
            coords = (int(dlg.x.text()), int(dlg.y.text()))
        if dlg.double.isChecked():
            double = True
        if dlg.wheel_dist.text() != '':
            wheel_dist = int(dlg.wheel_dist.text())
        self.current_elem_wrapper.click_input(
            button=button, coords=coords, double=double, wheel_dist=wheel_dist)
        self.__write_method('click_input(button=' + button + ', coords=' + str(coords) + ', double=' + str(double) + ', wheel_dist=' + str(wheel_dist) + ')', False)

    def __close(self):
        if self.current_elem_wrapper:
            if self.current_elem_wrapper.is_dialog():
                self.current_elem_wrapper.close()
                self.__write_method('close()', False)
                self.__refresh()

    def __descendants(self):
        dlg = InfoDialog(
            'descendants()', self.current_elem_wrapper.descendants(), self)
        dlg.exec()
        self.__write_method('descendants()', True)

    #def __draw_outline(self):
    #    self.current_elem_wrapper.draw_outline()

    def __set_focus(self):
        self.current_elem_wrapper.set_focus()
        self.__write_method('set_focus()', False)

    def __texts(self):
        dlg = InfoDialog('texts()', self.current_elem_wrapper.texts(), self)
        dlg.exec()
        self.__write_method('texts()', True)

    def __type_keys(self):
        # TODO show dialog to choose keys + write method
        self.current_elem_wrapper.type_keys()

    def __wait_enabled(self):
        self.current_elem_wrapper.wait_enabled()
        self.__write_method('wait_enabled()', False)

    def __wait_not_enabled(self):
        self.current_elem_wrapper.wait_not_enabled()
        self.__write_method('wait_not_enabled()', False)

    def __wait_not_visible(self):
        self.current_elem_wrapper.wait_not_visible()
        self.__write_method('wait_not_visible()', False)

    def __wait_visible(self):
        self.current_elem_wrapper.wait_visible()
        self.__write_method('wait_visible()', False)

    def __window_text(self):
        dlg = InfoDialog(
            'window_text', [self.current_elem_wrapper.window_text()], self)
        dlg.exec()
        self.__write_method('window_text()', True)

    # Hwnd Wrapper Methods

    # Hwnd Controls Wrappers Methods

    # UIA Wrapper Methods

    # UIA Controls Wrappers Methods


class MyTreeModel(QStandardItemModel):
    def __init__(self, element_info, backend):
        QStandardItemModel.__init__(self)
        root_node = self.invisibleRootItem()
        self.props_dict = {}
        self.info_dict = {}
        self.backend = backend
        self.branch = QStandardItem(self.__node_name(element_info))
        self.branch.setEditable(False)
        root_node.appendRow(self.branch)
        self.__generate_props_dict(element_info)
        self.__get_next(element_info, self.branch)

    def __get_next(self, element_info, parent):
        for child in element_info.children():
            self.__generate_props_dict(child)
            child_item \
                = QStandardItem(self.__node_name(child))
            child_item.setEditable(False)
            parent.appendRow(child_item)
            self.__get_next(child, child_item)

    def __node_name(self, element_info):
        if 'uia' == self.backend:
            return '%s "%s" (%s)' % (str(element_info.control_type),
                                     str(element_info.name),
                                     id(element_info))
        return '"%s" (%s)' % (str(element_info.name), id(element_info))

    def __generate_props_dict(self, element_info):
        props = [
            ['control_id', str(element_info.control_id)],
            ['class_name', str(element_info.class_name)],
            ['enabled', str(element_info.enabled)],
            ['handle', str(element_info.handle)],
            ['name', str(element_info.name)],
            ['process_id', str(element_info.process_id)],
            ['rectangle', str(element_info.rectangle)],
            ['rich_text', str(element_info.rich_text)],
            ['visible', str(element_info.visible)]
        ]

        props_win32 = [
        ] if (self.backend == 'win32') else []

        props_uia = [
            ['auto_id', str(element_info.auto_id)],
            ['control_type', str(element_info.control_type)],
            ['element', str(element_info.element)],
            ['framework_id', str(element_info.framework_id)],
            ['runtime_id', str(element_info.runtime_id)],
            ['access_key', str(element_info.access_key)],
            ['legacy_action', str(element_info.legacy_action)],
            ['legacy_descr', str(element_info.legacy_descr)],
            ['legacy_help', str(element_info.legacy_help)],
            ['legacy_name', str(element_info.legacy_name)],
            ['legacy_shortcut', str(element_info.legacy_shortcut)],
            ['legacy_value', str(element_info.legacy_value)],
            ['accelerator', str(element_info.accelerator)],
            ['value', str(element_info.value)],
            ['parent', str(element_info.parent)],
            ['top_level_parent', str(element_info.top_level_parent)]
        ] if (self.backend == 'uia') else []

        props.extend(props_uia)
        props.extend(props_win32)
        node_dict = {self.__node_name(element_info): props}
        self.props_dict.update(node_dict)
        self.info_dict.update({self.__node_name(element_info): element_info})


class MyTableModel(QAbstractTableModel):
    def __init__(self, datain, parent=None, *args):
        QAbstractTableModel.__init__(self, parent, *args)
        self.arraydata = datain
        self.header_labels = ['Property', 'Value']

    def rowCount(self, parent):
        return len(self.arraydata)

    def columnCount(self, parent):
        return len(self.arraydata[0])

    def data(self, index, role):
        if not index.isValid():
            return QVariant()
        elif role != Qt.DisplayRole:
            return QVariant()
        return QVariant(self.arraydata[index.row()][index.column()])

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            return self.header_labels[section]
        return QAbstractTableModel.headerData(self, section, orientation, role)


if __name__ == "__main__":
    main()
