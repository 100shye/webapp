import sys
import json
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QLineEdit, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QComboBox, QFileDialog, QInputDialog, QMessageBox
)
import pandas as pd

# --- Configuration ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
RECIPE_DATA_FILE = os.path.join(SCRIPT_DIR, "recipes.json")

class RecipeManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dynamic Recipe Manager")
        self.setGeometry(100, 100, 1200, 800)

        self.recipe_store = {}
        self.INGREDIENT_NAME_COLUMN = "Chamber"
        self.STEP_PROPERTY_COLUMN = "속성 (Property)"
        
        self.setup_ui()
        self.load_data_from_file()

    def setup_ui(self):
        self.recipe_selector = QComboBox()
        self.recipe_selector.currentIndexChanged.connect(self.display_selected_recipe)
        self.title_input = QLineEdit()
        self.title_input.setPlaceholderText("레시피 제목을 입력하세요")
        self.ingredient_table = QTableWidget(0, 2)
        self.ingredient_table.setHorizontalHeaderLabels([self.INGREDIENT_NAME_COLUMN, "Recipe"])
        self.step_table = QTableWidget(0, 1)
        self.step_table.setHorizontalHeaderLabels([self.STEP_PROPERTY_COLUMN])
        
        main_layout = QVBoxLayout()
        top_controls_layout = QHBoxLayout()
        top_controls_layout.addWidget(QLabel("레시피 선택:"))
        top_controls_layout.addWidget(self.recipe_selector)
        new_recipe_btn = QPushButton("새 레시피")
        new_recipe_btn.clicked.connect(self.clear_ui_for_new_recipe)
        top_controls_layout.addWidget(new_recipe_btn)
        delete_recipe_btn = QPushButton("레시피 삭제")
        delete_recipe_btn.clicked.connect(self.delete_recipe)
        top_controls_layout.addWidget(delete_recipe_btn)
        title_layout = QHBoxLayout()
        title_layout.addWidget(QLabel("레시피 제목:"))
        title_layout.addWidget(self.title_input)
        ing_btn_layout = QHBoxLayout()
        add_ing_row_btn = QPushButton("재료 행 추가")
        add_ing_row_btn.clicked.connect(self.add_ingredient_row)
        del_ing_row_btn = QPushButton("재료 행 삭제")
        del_ing_row_btn.clicked.connect(self.delete_ingredient_row)
        add_ing_col_btn = QPushButton("속성 열 추가")
        add_ing_col_btn.clicked.connect(self.add_ingredient_column)
        del_ing_col_btn = QPushButton("속성 열 삭제")
        del_ing_col_btn.clicked.connect(self.delete_ingredient_column)
        ing_btn_layout.addWidget(add_ing_row_btn)
        ing_btn_layout.addWidget(del_ing_row_btn)
        ing_btn_layout.addStretch()
        ing_btn_layout.addWidget(add_ing_col_btn)
        ing_btn_layout.addWidget(del_ing_col_btn)
        step_btn_layout = QHBoxLayout()
        add_step_row_btn = QPushButton("조리단계 행 추가")
        add_step_row_btn.clicked.connect(self.add_step_row)
        add_step_col_btn = QPushButton("조리단계 열 추가")
        add_step_col_btn.clicked.connect(self.add_step_column)
        del_step_row_btn = QPushButton("조리단계 행 삭제")
        del_step_row_btn.clicked.connect(self.delete_step_row)
        step_btn_layout.addStretch()
        step_btn_layout.addWidget(add_step_row_btn)
        step_btn_layout.addWidget(add_step_col_btn)
        step_btn_layout.addWidget(del_step_row_btn)
        save_layout = QHBoxLayout()
        save_recipe_btn = QPushButton("레시피 파일에 저장")
        save_recipe_btn.clicked.connect(self.save_current_recipe)
        export_excel_btn = QPushButton("엑셀로 내보내기")
        export_excel_btn.clicked.connect(self.export_to_excel)
        save_layout.addStretch()
        save_layout.addWidget(save_recipe_btn)
        save_layout.addWidget(export_excel_btn)
        
        main_layout.addLayout(top_controls_layout)
        main_layout.addLayout(title_layout)
        main_layout.addSpacing(20)
        main_layout.addWidget(QLabel("<h3>재료</h3>"))
        
        # --- MODIFICATION: Setting the stretch factor to 1 ---
        main_layout.addWidget(self.ingredient_table, 1)
        
        main_layout.addLayout(ing_btn_layout)
        main_layout.addSpacing(20)
        main_layout.addWidget(QLabel("<h3>조리 단계</h3>"))

        # --- MODIFICATION: Setting the stretch factor to 2 (twice as big) ---
        main_layout.addWidget(self.step_table, 2)
        
        main_layout.addLayout(step_btn_layout)
        main_layout.addLayout(save_layout)
        
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def display_selected_recipe(self):
        title = self.recipe_selector.currentText()
        if not title or title not in self.recipe_store:
            self.clear_ui_for_new_recipe(); return
        
        recipe_data = self.recipe_store[title]
        self.title_input.setText(title)
        
        self.ingredient_table.blockSignals(True)
        ingredients = recipe_data.get("ingredients", [])
        all_headers = set(); [all_headers.update(d.keys()) for d in ingredients]
        header_list = []
        if self.INGREDIENT_NAME_COLUMN in all_headers:
            header_list.append(self.INGREDIENT_NAME_COLUMN)
            all_headers.remove(self.INGREDIENT_NAME_COLUMN)
        header_list.extend(sorted(list(all_headers)))
        self.ingredient_table.setRowCount(0); self.ingredient_table.setColumnCount(len(header_list))
        self.ingredient_table.setHorizontalHeaderLabels(header_list)
        self.ingredient_table.setRowCount(len(ingredients))
        for r, d in enumerate(ingredients):
            for c, h in enumerate(header_list):
                self.ingredient_table.setItem(r, c, QTableWidgetItem(str(d.get(h, ""))))
        self.ingredient_table.blockSignals(False)
        
        self.update_step_property_options()
        
        self.step_table.setRowCount(0)
        step_data = recipe_data.get("steps", {})
        step_rows = step_data.get("rows", [])
        step_columns = step_data.get("columns", [self.STEP_PROPERTY_COLUMN])
        self.step_table.setColumnCount(len(step_columns))
        self.step_table.setHorizontalHeaderLabels(step_columns)
        self.step_table.setRowCount(len(step_rows))
        
        for row_idx, row_data in enumerate(step_rows):
            prop_combo = QComboBox()
            prop_combo.addItems(self.get_ingredient_headers())
            prop_value = row_data.get(self.STEP_PROPERTY_COLUMN, row_data.get("ingredient", ""))
            prop_combo.setCurrentText(prop_value)
            self.step_table.setCellWidget(row_idx, 0, prop_combo)
            
            for col_idx, col_name in enumerate(step_columns[1:], start=1):
                item = QTableWidgetItem(row_data.get(col_name, ""))
                self.step_table.setItem(row_idx, col_idx, item)

    def save_current_recipe(self):
        title = self.title_input.text().strip()
        if not title: QMessageBox.warning(self, "입력 오류", "레시피 제목을 입력하세요."); return

        ingredients = []
        headers = [self.ingredient_table.horizontalHeaderItem(i).text() for i in range(self.ingredient_table.columnCount())]
        for row in range(self.ingredient_table.rowCount()):
            ing_dict = {h: self.ingredient_table.item(row, c).text() for c, h in enumerate(headers) if self.ingredient_table.item(row, c) and self.ingredient_table.item(row, c).text().strip()}
            if ing_dict: ingredients.append(ing_dict)

        step_columns = [self.step_table.horizontalHeaderItem(i).text() for i in range(self.step_table.columnCount())]
        step_rows = []
        for row in range(self.step_table.rowCount()):
            row_data = {}
            prop_combo = self.step_table.cellWidget(row, 0)
            if prop_combo:
                row_data[self.STEP_PROPERTY_COLUMN] = prop_combo.currentText()
            
            for col_idx, col_name in enumerate(step_columns[1:], start=1):
                item = self.step_table.item(row, col_idx)
                if item: row_data[col_name] = item.text()
            step_rows.append(row_data)

        self.recipe_store[title] = { "ingredients": ingredients, "steps": {"columns": step_columns, "rows": step_rows} }
        self.save_data_to_file()
        self.update_recipe_selector()
        self.recipe_selector.setCurrentText(title)
        QMessageBox.information(self, "성공", f"'{title}' 레시피가 파일에 저장되었습니다.")

    def add_ingredient_row(self): self.ingredient_table.insertRow(self.ingredient_table.rowCount())
    def delete_ingredient_row(self):
        if self.ingredient_table.currentRow() > -1: self.ingredient_table.removeRow(self.ingredient_table.currentRow())
    def add_ingredient_column(self):
        text, ok = QInputDialog.getText(self, "새 속성 추가", "추가할 열의 이름을 입력하세요:")
        if ok and text.strip():
            if text in self.get_ingredient_headers():
                QMessageBox.warning(self, "오류", "같은 이름의 열이 이미 존재합니다."); return
            col_pos = self.ingredient_table.columnCount()
            self.ingredient_table.insertColumn(col_pos)
            self.ingredient_table.setHorizontalHeaderItem(col_pos, QTableWidgetItem(text))
            self.update_step_property_options()
    def delete_ingredient_column(self):
        col = self.ingredient_table.currentColumn()
        if col > -1:
            header = self.ingredient_table.horizontalHeaderItem(col).text()
            if header == self.INGREDIENT_NAME_COLUMN:
                QMessageBox.warning(self, "오류", f"기본 '{self.INGREDIENT_NAME_COLUMN}' 열은 삭제할 수 없습니다."); return
            self.ingredient_table.removeColumn(col)
            self.update_step_property_options()

    def get_ingredient_headers(self):
        return [self.ingredient_table.horizontalHeaderItem(i).text() for i in range(self.ingredient_table.columnCount())]

    def add_step_row(self):
        row = self.step_table.rowCount()
        self.step_table.insertRow(row)
        prop_combo = QComboBox()
        prop_combo.addItems(self.get_ingredient_headers())
        self.step_table.setCellWidget(row, 0, prop_combo)
    def delete_step_row(self):
        if self.step_table.currentRow() > -1: self.step_table.removeRow(self.step_table.currentRow())
    def add_step_column(self):
        col = self.step_table.columnCount()
        step_num = 1
        while f"Step {step_num}" in [self.step_table.horizontalHeaderItem(i).text() for i in range(1, col)]: step_num += 1
        self.step_table.insertColumn(col)
        self.step_table.setHorizontalHeaderItem(col, QTableWidgetItem(f"Step {step_num}"))
        
    def update_step_property_options(self):
        property_names = self.get_ingredient_headers()
        for row in range(self.step_table.rowCount()):
            combo = self.step_table.cellWidget(row, 0)
            if isinstance(combo, QComboBox):
                current_text = combo.currentText()
                combo.blockSignals(True)
                combo.clear()
                combo.addItems(property_names)
                if current_text in property_names:
                    combo.setCurrentText(current_text)
                combo.blockSignals(False)
    
    def load_data_from_file(self):
        if not os.path.exists(RECIPE_DATA_FILE): self.recipe_store = {}; return
        try:
            with open(RECIPE_DATA_FILE, 'r', encoding='utf-8') as f: self.recipe_store = json.load(f)
            self.update_recipe_selector()
        except (json.JSONDecodeError, FileNotFoundError): self.recipe_store = {}; QMessageBox.warning(self, "Load Error", f"Could not load {RECIPE_DATA_FILE}. Starting fresh.")
    def save_data_to_file(self):
        try:
            with open(RECIPE_DATA_FILE, 'w', encoding='utf-8') as f: json.dump(self.recipe_store, f, indent=4, ensure_ascii=False)
        except Exception as e: QMessageBox.critical(self, "Save Error", f"Could not save to {RECIPE_DATA_FILE}: {e}")
    def update_recipe_selector(self):
        current_selection = self.recipe_selector.currentText(); self.recipe_selector.blockSignals(True)
        self.recipe_selector.clear(); self.recipe_selector.addItems(sorted(self.recipe_store.keys()))
        self.recipe_selector.setCurrentText(current_selection); self.recipe_selector.blockSignals(False)
        if self.recipe_selector.currentIndex() == -1 and self.recipe_selector.count() > 0: self.recipe_selector.setCurrentIndex(0)
        elif self.recipe_selector.count() == 0: self.clear_ui_for_new_recipe()
    def delete_recipe(self):
        title = self.recipe_selector.currentText();
        if not title: return
        reply = QMessageBox.question(self, "삭제 확인", f"'{title}' 레시피를 정말 삭제하시겠습니까?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes and title in self.recipe_store:
            del self.recipe_store[title]; self.save_data_to_file(); self.update_recipe_selector()
            QMessageBox.information(self, "성공", f"'{title}' 레시피가 삭제되었습니다.")
    def export_to_excel(self):
        title = self.title_input.text().strip();
        if not title or title not in self.recipe_store: return
        path, _ = QFileDialog.getSaveFileName(self, "엑셀로 내보내기", f"{title}.xlsx", "Excel Files (*.xlsx)")
        if not path: return
        recipe_data = self.recipe_store[title]
        ingredients_df = pd.DataFrame(recipe_data.get("ingredients", [])); steps_df = pd.DataFrame(recipe_data.get("steps", {}).get("rows", []))
        try:
            with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
                ingredients_df.to_excel(writer, sheet_name="Ingredients", index=False)
                steps_df.to_excel(writer, sheet_name="Cooking Steps", index=False)
            QMessageBox.information(self, "내보내기 성공", f"레시피를 {path}에 성공적으로 저장했습니다.")
        except Exception as e: QMessageBox.critical(self, "내보내기 실패", f"내보내기 중 오류 발생: {e}")
    def clear_ui_for_new_recipe(self):
        self.title_input.clear()
        self.ingredient_table.setRowCount(0); self.ingredient_table.setColumnCount(2)
        self.ingredient_table.setHorizontalHeaderLabels([self.INGREDIENT_NAME_COLUMN, "Recipe"])
        self.step_table.setRowCount(0); self.step_table.setColumnCount(1)
        self.step_table.setHorizontalHeaderLabels([self.STEP_PROPERTY_COLUMN])
        self.recipe_selector.setCurrentIndex(-1)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = RecipeManager()
    window.show()
    sys.exit(app.exec_())
