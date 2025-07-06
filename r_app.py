import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QLineEdit, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QComboBox, QFileDialog, QInputDialog, QMessageBox
)
import pandas as pd
import xlsxwriter

class RecipeManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("레시피 관리자")
        self.setGeometry(100, 100, 1200, 800)

        self.recipe_store = {}
        self.step_count = 0

        self.recipe_selector = QComboBox()
        self.recipe_selector.currentIndexChanged.connect(self.load_selected_recipe)

        self.title_input = QLineEdit()

        self.ingredient_table = QTableWidget(0, 3)
        self.ingredient_table.setHorizontalHeaderLabels(["재료", "분량", "단위"])

        self.step_table = QTableWidget(0, 1)
        self.step_table.setHorizontalHeaderLabels(["재료명"])
        self.step_columns = ["재료명"]

        self.setup_ui()

    def setup_ui(self):
        main_layout = QVBoxLayout()

        # 상단: 레시피 제목 + 저장/불러오기
        selector_layout = QHBoxLayout()
        selector_layout.addWidget(QLabel("레시피 목록"))
        selector_layout.addWidget(self.recipe_selector)

        selector_btns = QHBoxLayout()
        load_btn = QPushButton("불러오기")
        load_btn.clicked.connect(self.load_selected_recipe)
        new_btn = QPushButton("새 레시피")
        new_btn.clicked.connect(self.new_recipe)
        selector_btns.addWidget(load_btn)
        selector_btns.addWidget(new_btn)

        title_layout = QHBoxLayout()
        title_layout.addWidget(QLabel("레시피 제목"))
        title_layout.addWidget(self.title_input)

        # 재료 테이블 + 버튼
        ing_btn_layout = QHBoxLayout()
        add_ing = QPushButton("재료 추가")
        add_ing.clicked.connect(self.add_ingredient_row)
        del_ing = QPushButton("재료 삭제")
        del_ing.clicked.connect(self.delete_ingredient_row)
        ing_btn_layout.addWidget(add_ing)
        ing_btn_layout.addWidget(del_ing)

        # 조리 단계 테이블 + 버튼
        step_btn_layout = QHBoxLayout()
        add_row = QPushButton("조리단계 재료 행 추가")
        add_row.clicked.connect(self.add_step_row)
        add_col = QPushButton("스텝 추가")
        add_col.clicked.connect(self.add_step_column)
        del_row = QPushButton("조리단계 행 삭제")
        del_row.clicked.connect(self.delete_step_row)
        step_btn_layout.addWidget(add_row)
        step_btn_layout.addWidget(add_col)
        step_btn_layout.addWidget(del_row)

        # 저장/엑셀 버튼
        save_btn = QPushButton("CSV로 저장")
        save_btn.clicked.connect(self.save_recipe)

        excel_btn = QPushButton("엑셀로 저장")
        excel_btn.clicked.connect(self.export_to_excel)

        save_layout = QHBoxLayout()
        save_layout.addWidget(save_btn)
        save_layout.addWidget(excel_btn)

        main_layout.addLayout(selector_layout)
        main_layout.addLayout(selector_btns)
        main_layout.addLayout(title_layout)
        main_layout.addWidget(QLabel("재료 목록"))
        main_layout.addWidget(self.ingredient_table)
        main_layout.addLayout(ing_btn_layout)
        main_layout.addWidget(QLabel("조리 단계 목록"))
        main_layout.addWidget(self.step_table)
        main_layout.addLayout(step_btn_layout)
        main_layout.addLayout(save_layout)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def add_ingredient_row(self):
        row = self.ingredient_table.rowCount()
        self.ingredient_table.insertRow(row)
        self.ingredient_table.setItem(row, 0, QTableWidgetItem(""))
        
        combo = QComboBox()
        options = ["100g", "200g", "1컵", "직접 입력..."]
        combo.addItems(options)
        combo.currentIndexChanged.connect(lambda idx, r=row: self.handle_custom_quantity(r))
        self.ingredient_table.setCellWidget(row, 1, combo)
        self.ingredient_table.setItem(row, 2, QTableWidgetItem(""))

        self.update_step_reagent_options()

    def handle_custom_quantity(self, row):
        combo = self.ingredient_table.cellWidget(row, 1)
        if combo.currentText() == "직접 입력...":
            text, ok = QInputDialog.getText(self, "직접 입력", "분량을 입력하세요:")
            if ok and text:
                if text not in [combo.itemText(i) for i in range(combo.count())]:
                    combo.insertItem(combo.count() - 1, text)
                combo.setCurrentText(text)

    def delete_ingredient_row(self):
        row = self.ingredient_table.currentRow()
        if row >= 0:
            self.ingredient_table.removeRow(row)
            self.update_step_reagent_options()

    def update_step_reagent_options(self):
        reagents = []
        for row in range(self.ingredient_table.rowCount()):
            item = self.ingredient_table.item(row, 0)
            if item and item.text().strip():
                reagents.append(item.text().strip())

        for row in range(self.step_table.rowCount()):
            combo = self.step_table.cellWidget(row, 0)
            if isinstance(combo, QComboBox):
                current = combo.currentText()
                combo.blockSignals(True)
                combo.clear()
                combo.addItems(reagents)
                combo.setCurrentText(current if current in reagents else "")
                combo.blockSignals(False)

    def add_step_row(self):
        row = self.step_table.rowCount()
        self.step_table.insertRow(row)

        # 재료명 콤보박스
        combo = QComboBox()
        reagents = [self.ingredient_table.item(r, 0).text().strip() for r in range(self.ingredient_table.rowCount()) if self.ingredient_table.item(r, 0)]
        combo.addItems(reagents)
        combo.currentIndexChanged.connect(lambda idx, r=row: self.update_step_quantity_cells(r))
        self.step_table.setCellWidget(row, 0, combo)

        for col in range(1, self.step_table.columnCount()):
            self.step_table.setCellWidget(row, col, QComboBox())

    def delete_step_row(self):
        row = self.step_table.currentRow()
        if row >= 0:
            self.step_table.removeRow(row)

    def add_step_column(self):
        self.step_count += 1
        step_name = f"Step{self.step_count}"
        col = self.step_table.columnCount()
        self.step_table.insertColumn(col)
        self.step_table.setHorizontalHeaderItem(col, QTableWidgetItem(step_name))
        self.step_columns.append(step_name)

        for row in range(self.step_table.rowCount()):
            self.step_table.setCellWidget(row, col, QComboBox())

    def update_step_quantity_cells(self, row):
        combo = self.step_table.cellWidget(row, 0)
        if not combo:
            return
        selected = combo.currentText()
        quantity = ""
        for r in range(self.ingredient_table.rowCount()):
            name_item = self.ingredient_table.item(r, 0)
            if name_item and name_item.text().strip() == selected:
                q_combo = self.ingredient_table.cellWidget(r, 1)
                if q_combo:
                    quantity = q_combo.currentText()
                break

        for col in range(1, self.step_table.columnCount()):
            step_cell = self.step_table.cellWidget(row, col)
            if isinstance(step_cell, QComboBox):
                step_cell.clear()
                if quantity:
                    step_cell.addItems([quantity, "직접 입력..."])
                else:
                    step_cell.addItem("직접 입력...")

    def save_recipe(self):
        name = self.title_input.text().strip()
        if not name:
            QMessageBox.warning(self, "알림", "레시피 제목을 입력하세요.")
            return

        ingredients = []
        for row in range(self.ingredient_table.rowCount()):
            name_item = self.ingredient_table.item(row, 0)
            quantity_combo = self.ingredient_table.cellWidget(row, 1)
            unit_item = self.ingredient_table.item(row, 2)
            ingredients.append({
                "재료": name_item.text() if name_item else "",
                "분량": quantity_combo.currentText() if quantity_combo else "",
                "단위": unit_item.text() if unit_item else ""
            })

        steps = []
        for row in range(self.step_table.rowCount()):
            step_row = {}
            combo = self.step_table.cellWidget(row, 0)
            step_row["재료명"] = combo.currentText() if combo else ""
            for col in range(1, self.step_table.columnCount()):
                cell = self.step_table.cellWidget(row, col)
                value = cell.currentText() if isinstance(cell, QComboBox) else ""
                step_row[self.step_columns[col]] = value
            steps.append(step_row)

        self.recipe_store[name] = {
            "ingredients": ingredients,
            "steps": steps
        }

        self.recipe_selector.clear()
        self.recipe_selector.addItems(self.recipe_store.keys())
        QMessageBox.information(self, "저장 완료", f"{name} 레시피가 저장되었습니다.")

    def load_selected_recipe(self):
        name = self.recipe_selector.currentText()
        if not name or name not in self.recipe_store:
            return

        data = self.recipe_store[name]
        self.title_input.setText(name)

        # 재료 테이블
        self.ingredient_table.setRowCount(0)
        for ing in data["ingredients"]:
            self.add_ingredient_row()
            r = self.ingredient_table.rowCount() - 1
            self.ingredient_table.setItem(r, 0, QTableWidgetItem(ing["재료"]))
            combo = self.ingredient_table.cellWidget(r, 1)
            combo.setCurrentText(ing["분량"])
            self.ingredient_table.setItem(r, 2, QTableWidgetItem(ing["단위"]))

        # 단계 테이블
        self.step_table.setRowCount(0)
        self.step_table.setColumnCount(1)
        self.step_table.setHorizontalHeaderItem(0, QTableWidgetItem("재료명"))
        self.step_columns = ["재료명"]
        self.step_count = 0

        for step in data["steps"]:
            self.add_step_column()
        for step in data["steps"]:
            self.add_step_row()
            r = self.step_table.rowCount() - 1
            self.step_table.cellWidget(r, 0).setCurrentText(step["재료명"])
            for c, col_name in enumerate(self.step_columns[1:], start=1):
                cell = self.step_table.cellWidget(r, c)
                cell.setCurrentText(step.get(col_name, ""))

    def new_recipe(self):
        self.title_input.clear()
        self.ingredient_table.setRowCount(0)
        self.step_table.setRowCount(0)
        self.step_table.setColumnCount(1)
        self.step_table.setHorizontalHeaderItem(0, QTableWidgetItem("재료명"))
        self.step_columns = ["재료명"]
        self.step_count = 0

    def export_to_excel(self):
        name = self.title_input.text().strip()
        if not name:
            QMessageBox.warning(self, "알림", "레시피 제목을 입력하세요.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "엑셀로 저장", f"{name}.xlsx", "Excel Files (*.xlsx)")
        if not path:
            return

        ingredients = pd.DataFrame(self.recipe_store[name]["ingredients"])
        steps = pd.DataFrame(self.recipe_store[name]["steps"])

        with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
            ingredients.to_excel(writer, sheet_name="재료", index=False)
            steps.to_excel(writer, sheet_name="조리단계", index=False)
        QMessageBox.information(self, "엑셀 저장", f"{path}로 저장되었습니다.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = RecipeManager()
    window.show()
    sys.exit(app.exec_())
