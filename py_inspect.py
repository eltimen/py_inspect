from PyQt5.QtCore import QLocale, QSortFilterProxyModel
from PyQt5.QtCore import QCoreApplication
from PyQt5.QtCore import QSettings
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QAbstractTableModel
from PyQt5.QtCore import QVariant
from PyQt5.QtGui import QStandardItemModel, QValidator
from PyQt5.QtGui import QStandardItem
from PyQt5.QtWidgets import QApplication, QCompleter
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QGridLayout
from PyQt5.QtWidgets import QHBoxLayout
from PyQt5.QtWidgets import QLabel
from PyQt5.QtWidgets import QComboBox
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QTreeView
from PyQt5.QtWidgets import QTableView
import psutil
import re
import sys
import warnings

warnings.simplefilter("ignore", UserWarning)
sys.coinit_flags = 2
from pywinauto import backend, Application


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    w = MyWindow()
    w.show()
    sys.exit(app.exec_())


class MyWindow(QWidget):
    def __init__(self):
        super(MyWindow, self).__init__()

        self.setMinimumSize(1000, 1000)
        self.setLocale(QLocale(QLocale.English, QLocale.UnitedStates))
        self.setWindowTitle(
            QCoreApplication.translate("MainWindow", "PyInspect"))

        self.settings = QSettings('py_inspect', 'MainWindow')

        # Main layout
        self.mainLayout = QGridLayout()

        # Backend label
        self.backendLabel = QLabel("Backend Type")

        # Backend combobox
        self.comboBox = QComboBox()
        self.comboBox.setMouseTracking(False)
        self.comboBox.setMaxVisibleItems(5)
        self.comboBox.setObjectName("comboBox")

        for _backend in backend.registry.backends.keys():
            self.comboBox.addItem(_backend)

        # Add top widgets to main window
        self.mainLayout.addWidget(self.backendLabel, 0, 0, 1, 1)
        self.mainLayout.addWidget(self.comboBox, 0, 1, 1, 1)

        self.processComboBox = FilterComboBox()
        self.processComboBox.setEnabled(False)
        self.processComboBox.setMaxVisibleItems(30)

        self.__init_process_list()

        self.processRefreshButton = QPushButton("Refresh")
        self.processRefreshButton.setEnabled(False)
        self.processRefreshButton.clicked.connect(self.__init_process_list)

        self.injectButton = QPushButton("Inject!")
        self.injectButton.setEnabled(False)
        self.injectButton.clicked.connect(self.__show_process_tree)

        self.current_application = None

        # Add widgets related to process-specific backends
        self.processSelectorLayout = QHBoxLayout()
        self.processSelectorLayout.addWidget(self.processComboBox)
        self.processSelectorLayout.addWidget(self.processRefreshButton)
        self.mainLayout.addLayout(self.processSelectorLayout, 1, 0, 1, 1)
        self.mainLayout.addWidget(self.injectButton, 1, 1, 1, 1)

        self.tree_view = QTreeView()
        self.tree_view.setColumnWidth(0, 150)

        self.comboBox.setCurrentText('uia')
        self.__initialize_calc()

        self.table_view = QTableView()

        self.comboBox.activated[str].connect(self.__show_tree)

        # Add center widgets to main window
        self.mainLayout.addWidget(self.tree_view, 2, 0, 1, 1)
        self.mainLayout.addWidget(self.table_view, 2, 1, 1, 1)

        self.setLayout(self.mainLayout)
        geometry = self.settings.value('Geometry', bytes('', 'utf-8'))
        self.restoreGeometry(geometry)

    def __init_process_list(self):
        process_list = []
        for proc in psutil.process_iter():
            process_string = '{} ({})'.format(proc.name(), proc.pid)
            process_list.append(process_string)
            self.processComboBox.addItem(process_string, proc.pid)

    def __show_process_tree(self):
        _pid = self.processComboBox.currentData()
        _backend = self.comboBox.currentText()

        self.current_application = Application(backend=_backend).connect(pid=_pid)

        self.element_info \
            = backend.registry.backends[_backend].element_info_class(pid=_pid)
        self.tree_model = MyTreeModel(self.element_info, _backend)
        self.tree_model.setHeaderData(0, Qt.Horizontal, 'Controls')
        self.tree_view.setModel(self.tree_model)
        self.tree_view.clicked.connect(self.__show_property)

    def __initialize_calc(self, _backend='uia'):
        self.element_info \
            = backend.registry.backends[_backend].element_info_class()
        self.tree_model = MyTreeModel(self.element_info, _backend)
        self.tree_model.setHeaderData(0, Qt.Horizontal, 'Controls')
        self.tree_view.setModel(self.tree_model)
        self.tree_view.clicked.connect(self.__show_property)

    def __show_tree(self, text):
        backend = text
        if backend == 'wpf':
            self.processRefreshButton.setEnabled(True)
            self.processComboBox.setEnabled(True)
            self.injectButton.setEnabled(True)

            self.tree_view.setModel(None)
            self.table_view.setModel(None)
        else:
            self.processRefreshButton.setEnabled(False)
            self.processComboBox.setEnabled(False)
            self.injectButton.setEnabled(False)

            self.__initialize_calc(backend)

    def __show_property(self, index=None):
        data = index.data()
        self.table_model \
            = MyTableModel(self.tree_model.props_dict.get(data), self)
        self.table_view.wordWrap()
        self.table_view.setModel(self.table_model)
        self.table_view.setColumnWidth(1, 320)

    def closeEvent(self, event):
        geometry = self.saveGeometry()
        self.settings.setValue('Geometry', geometry)
        super(MyWindow, self).closeEvent(event)


class MyTreeModel(QStandardItemModel):
    def __init__(self, element_info, backend):
        QStandardItemModel.__init__(self)
        root_node = self.invisibleRootItem()
        self.props_dict = {}
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
                        ['runtime_id', str(element_info.runtime_id)]
                    ] if (self.backend == 'uia') else []

        props_wpf = [
            ['auto_id', str(element_info.auto_id)],
            ['control_type', str(element_info.control_type)],
            ['parent', str(element_info.parent)],
            ['framework_id', str(element_info.framework_id)],
            ['runtime_id', str(element_info.runtime_id)]
        ] if (self.backend == 'wpf') else []

        props.extend(props_wpf)
        props.extend(props_uia)
        props.extend(props_win32)
        node_dict = {self.__node_name(element_info): props}
        self.props_dict.update(node_dict)


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


class ComboBoxValidator(QValidator):
    def validate(self, input, pos):
        if input == '':
            return QValidator.Intermediate, input, pos
        combobox = self.parent()
        for i in range(combobox.count()):
            if combobox.itemText(i) == input:
                return QValidator.Acceptable, input, pos
            elif bool(re.match('^'+re.escape(input), combobox.itemText(i), re.I)):
                return QValidator.Intermediate, input, pos
        return QValidator.Invalid, input, pos


class FilterComboBox(QComboBox):
    # based on https://stackoverflow.com/a/50639066/
    def __init__(self, parent=None):
        super(FilterComboBox, self).__init__(parent)

        self.setFocusPolicy(Qt.StrongFocus)
        self.setEditable(True)

        # add a filter model to filter matching items
        self.pFilterModel = QSortFilterProxyModel(self)
        self.pFilterModel.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.pFilterModel.setSourceModel(self.model())

        # add a completer, which uses the filter model
        self.completer = QCompleter(self.pFilterModel, self)
        # always show all (filtered) completions
        self.completer.setCompletionMode(QCompleter.UnfilteredPopupCompletion)
        self.setCompleter(self.completer)

        # connect signals
        self.lineEdit().textEdited.connect(self.pFilterModel.setFilterFixedString)
        self.completer.activated.connect(self.on_completer_activated)

    # on selection of an item from the completer, select the corresponding item from combobox
    def on_completer_activated(self, text):
        if text:
            index = self.findText(text)
            self.setCurrentIndex(index)
            self.activated[str].emit(self.itemText(index))


if __name__ == "__main__":
    main()
