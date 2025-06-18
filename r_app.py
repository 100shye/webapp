
'''
pip install pyinstaller
pyinstaller --onefile --windowed r_manager.py
pause
'''


from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QTableWidget, QTableWidgetItem, QComboBox, QPushButton,
    QFileDialog, QInputDialog, QComboBox, QMessageBox
)
import sys
import pandas as pd

class RecipeManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("레시피 관리 앱")
        self.setGeometry(100, 100, 1100, 850)

        self.ingredients_list = ["강력분", "박력분", "중력분", "아몬드 가루", "직접 입력..."]
        self.recipe_store = {}
        self.existing_step_columns = set(["Step 번호", "설명"])

        self.title_input = QLineEdit()
        title_layout = QHBoxLayout()
        title_layout.addWidget(QLabel("레시피 제목:"))
        title_layout.addWidget(self.title_input)

        # 레시피 목록 콤보박스
        self.recipe_selector = QComboBox()
        self.load_button = QPushButton("불러오기")
        self.load_button.clicked.connect(self.load_selected_recipe)
        self.new_button = QPushButton("새 레시피")
        self.new_button.clicked.connect(self.new_recipe)

        selector_layout = QHBoxLayout()
        selector_layout.addWidget(QLabel("레시피 목록:"))
        selector_layout.addWidget(self.recipe_selector)
        selector_layout.addWidget(self.load_button)
        selector_layout.addWidget(self.new_button)

        # 재료 테이블
        self.ingredient_table = QTableWidget(0, 3)
        self.ingredient_table.setHorizontalHeaderLabels(["재료", "분량", "단위"])
        ingredient_btn = QPushButton("재료 추가")
        ingredient_btn.clicked.connect(self.add_ingredient_row)

        # 조리 단계 테이블
        self.step_table = QTableWidget(0, 2)
        self.step_table.setHorizontalHeaderLabels(["Step 번호", "설명"])
        step_row_btn = QPushButton("조리 단계 추가")
        step_row_btn.clicked.connect(self.add_step_row)
        step_row_delete_btn = QPushButton("조리 단계 삭제")
        step_row_delete_btn.clicked.connect(self.delete_step_row)
        step_col_btn = QPushButton("조리 단계 컬럼 추가")
        step_col_btn.clicked.connect(self.add_step_column)
        step_col_delete_btn = QPushButton("조리 단계 컬럼 삭제")
        step_col_delete_btn.clicked.connect(self.delete_step_column)

        # 저장 버튼
        save_btn = QPushButton("레시피 저장")
        save_btn.clicked.connect(self.save_recipe)

        save_excel_btn = QPushButton("엑셀로 저장")
        save_excel_btn.clicked.connect(self.export_to_excel)

        layout = QVBoxLayout()
        layout.addLayout(selector_layout)
        layout.addLayout(title_layout)

        layout.addWidget(QLabel("재료 목록"))
        layout.addWidget(self.ingredient_table)
        layout.addWidget(ingredient_btn)

        layout.addWidget(QLabel("조리 단계 목록"))
        layout.addWidget(self.step_table)
        step_btn_layout = QHBoxLayout()
        step_btn_layout.addWidget(step_row_btn)
        step_btn_layout.addWidget(step_row_delete_btn)
        step_btn_layout.addWidget(step_col_btn)
        step_btn_layout.addWidget(step_col_delete_btn)
        layout.addLayout(step_btn_layout)

        layout.addWidget(save_btn)
        layout.addWidget(save_excel_btn)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def add_ingredient_row(self):
        row = self.ingredient_table.rowCount()
        self.ingredient_table.insertRow(row)
        combo = QComboBox()
        combo.addItems(self.ingredients_list)
        combo.currentIndexChanged.connect(lambda idx, r=row: self.handle_custom_ingredient(r))
        self.ingredient_table.setCellWidget(row, 0, combo)
        combo.currentIndexChanged.connect(lambda: self.sync_step_columns_from_ingredients())

    def handle_custom_ingredient(self, row):
        combo = self.ingredient_table.cellWidget(row, 0)
        if combo.currentText() == "직접 입력...":
            text, ok = QInputDialog.getText(self, "재료 직접 입력", "새로운 재료명을 입력하세요:")
            if ok and text:
                if text not in self.ingredients_list:
                    self.ingredients_list.insert(-1, text)
                self.update_all_ingredient_combos(text)
                self.sync_step_columns_from_ingredients()

    def update_all_ingredient_combos(self, new_item):
        for row in range(self.ingredient_table.rowCount()):
            combo = self.ingredient_table.cellWidget(row, 0)
            current = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(self.ingredients_list)
            combo.setCurrentText(new_item if current == "직접 입력..." else current)
            combo.blockSignals(False)

    def sync_step_columns_from_ingredients(self):
        ingredients = set()
        for row in range(self.ingredient_table.rowCount()):
            combo = self.ingredient_table.cellWidget(row, 0)
            if combo:
                value = combo.currentText()
                if value and value not in self.existing_step_columns:
                    ingredients.add(value)

        for ingredient in ingredients:
            if ingredient not in self.existing_step_columns:
                col_index = self.step_table.columnCount()
                self.step_table.insertColumn(col_index)
                self.step_table.setHorizontalHeaderItem(col_index, QTableWidgetItem(ingredient))
                for row in range(self.step_table.rowCount()):
                    self.step_table.setItem(row, col_index, QTableWidgetItem(""))
                self.existing_step_columns.add(ingredient)

    def delete_step_column(self):
        text, ok = QInputDialog.getText(self, "컬럼 삭제", "삭제할 컬럼명을 입력하세요:")
        if ok and text and text in self.existing_step_columns and text not in {"Step 번호", "설명"}:
            col_to_delete = -1
            for i in range(self.step_table.columnCount()):
                if self.step_table.horizontalHeaderItem(i).text() == text:
                    col_to_delete = i
                    break
            if col_to_delete != -1:
                self.step_table.removeColumn(col_to_delete)
                self.existing_step_columns.remove(text)

    def delete_step_row(self):
        row = self.step_table.currentRow()
        if row >= 0:
            self.step_table.removeRow(row)
        else:
            QMessageBox.information(self, "행 삭제", "삭제할 행을 선택해주세요.")

    def add_step_row(self):
        row = self.step_table.rowCount()
        self.step_table.insertRow(row)
        self.step_table.setItem(row, 0, QTableWidgetItem(str(row + 1)))
        for col in range(1, self.step_table.columnCount()):
            self.step_table.setItem(row, col, QTableWidgetItem(""))

    def add_step_column(self):
        text, ok = QInputDialog.getText(self, "새 컬럼 추가", "새 컬럼명을 입력하세요:")
        if ok and text:
            col_index = self.step_table.columnCount()
            self.step_table.insertColumn(col_index)
            self.step_table.setHorizontalHeaderItem(col_index, QTableWidgetItem(text))
            for row in range(self.step_table.rowCount()):
                self.step_table.setItem(row, col_index, QTableWidgetItem(""))
            self.existing_step_columns.add(text)

    def save_recipe(self):
        recipe_name = self.title_input.text().strip()
        if not recipe_name:
            return

        ingredients = []
        for row in range(self.ingredient_table.rowCount()):
            combo = self.ingredient_table.cellWidget(row, 0)
            name = combo.currentText() if combo else ""
            amount = self.ingredient_table.item(row, 1)
            unit = self.ingredient_table.item(row, 2)
            ingredients.append({
                "재료": name,
                "분량": amount.text() if amount else "",
                "단위": unit.text() if unit else ""
            })

        steps = []
        headers = [self.step_table.horizontalHeaderItem(i).text() for i in range(self.step_table.columnCount())]
        for row in range(self.step_table.rowCount()):
            step_entry = {}
            for col in range(self.step_table.columnCount()):
                item = self.step_table.item(row, col)
                step_entry[headers[col]] = item.text() if item else ""
            steps.append(step_entry)

        self.recipe_store[recipe_name] = {
            "ingredients": ingredients,
            "steps": steps,
            "columns": headers
        }
        self.recipe_selector.clear()
        self.recipe_selector.addItems(self.recipe_store.keys())

    def load_selected_recipe(self):
        name = self.recipe_selector.currentText()
        if not name or name not in self.recipe_store:
            return

        data = self.recipe_store[name]
        self.title_input.setText(name)

        self.ingredient_table.setRowCount(0)
        for ing in data["ingredients"]:
            self.add_ingredient_row()
            row = self.ingredient_table.rowCount() - 1
            combo = self.ingredient_table.cellWidget(row, 0)
            combo.setCurrentText(ing["재료"])
            self.ingredient_table.setItem(row, 1, QTableWidgetItem(ing["분량"]))
            self.ingredient_table.setItem(row, 2, QTableWidgetItem(ing["단위"]))

        self.step_table.setColumnCount(len(data["columns"]))
        self.step_table.setHorizontalHeaderLabels(data["columns"])
        self.step_table.setRowCount(len(data["steps"]))
        self.existing_step_columns = set(data["columns"])
        for row, step in enumerate(data["steps"]):
            for col, key in enumerate(data["columns"]):
                self.step_table.setItem(row, col, QTableWidgetItem(step.get(key, "")))

    def new_recipe(self):
        self.title_input.clear()
        self.ingredient_table.setRowCount(0)
        self.step_table.setRowCount(0)
        self.step_table.setColumnCount(2)
        self.step_table.setHorizontalHeaderLabels(["Step 번호", "설명"])
        self.existing_step_columns = set(["Step 번호", "설명"])

    def export_to_excel(self):
        recipe_name = self.title_input.text().strip()
        if not recipe_name:
            return

        ingredients = []
        for row in range(self.ingredient_table.rowCount()):
            combo = self.ingredient_table.cellWidget(row, 0)
            name = combo.currentText() if combo else ""
            amount = self.ingredient_table.item(row, 1)
            unit = self.ingredient_table.item(row, 2)
            ingredients.append({
                "재료": name,
                "분량": amount.text() if amount else "",
                "단위": unit.text() if unit else ""
            })

        steps = []
        headers = [self.step_table.horizontalHeaderItem(i).text() for i in range(self.step_table.columnCount())]
        for row in range(self.step_table.rowCount()):
            step_entry = {}
            for col in range(self.step_table.columnCount()):
                item = self.step_table.item(row, col)
                step_entry[headers[col]] = item.text() if item else ""
            steps.append(step_entry)

        df_ingredients = pd.DataFrame(ingredients)
        df_steps = pd.DataFrame(steps)

        path, _ = QFileDialog.getSaveFileName(self, "엑셀로 저장", f"{recipe_name}.xlsx", "Excel Files (*.xlsx)")
        if path:
            with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
                df_ingredients.to_excel(writer, sheet_name='재료', index=False)
                df_steps.to_excel(writer, sheet_name='조리단계', index=False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = RecipeManager()
    window.show()
    sys.exit(app.exec_())