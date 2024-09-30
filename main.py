import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel, QTabWidget,
    QLineEdit, QComboBox, QDateEdit, QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
    QHBoxLayout, QToolBar, QAction, QFileDialog
)
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import Qt, QDate
import matplotlib.pyplot as plt
import numpy as np
from qt_material import apply_stylesheet
import pandas as pd
from docx import Document

class RoleManager:
    def __init__(self):
        self.user_role = "Admin"  # Rôle par défaut : Admin

    def is_admin(self):
        return self.user_role == "Admin"

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Cheetah Cost - Gestion des Coûts")
        self.setGeometry(100, 100, 1200, 800)
        self.role_manager = RoleManager()
        self.init_ui()

    def init_ui(self):
        self.welcome_screen = QWidget()
        self.setCentralWidget(self.welcome_screen)
        layout = QVBoxLayout()

        label = QLabel("Bienvenue sur Cheetah Cost")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 24px; font-weight: bold; color: #3E4A59;")
        layout.addWidget(label)

        start_button = QPushButton("Commencer", self)
        start_button.clicked.connect(self.open_project_creation)
        start_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 18px;
                font-weight: bold;
                padding: 10px;
                border-radius: 8px;
            }
            QPushButton:hover {
                background-color: #388E3C;
            }
        """)
        layout.addWidget(start_button, alignment=Qt.AlignCenter)

        self.welcome_screen.setLayout(layout)

    def open_project_creation(self):
        self.project_creation_screen = ProjectCreationScreen(self)
        self.setCentralWidget(self.project_creation_screen)

    def open_project_analysis_screen(self):
        self.project_analysis_screen = ProjectAnalysisScreen(self)
        self.setCentralWidget(self.project_analysis_screen)

    def return_to_main_menu(self):
        self.init_ui()

class ProjectCreationScreen(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        self.tab_widget = QTabWidget()

        # Onglet Période de Base
        base_period_tab = QWidget()
        base_period_layout = QVBoxLayout()

        self.project_name_input = QLineEdit()
        self.project_desc_input = QLineEdit()
        self.start_date_input = QDateEdit()
        self.start_date_input.setDate(QDate.currentDate())
        self.end_date_input = QDateEdit()
        self.end_date_input.setDate(QDate.currentDate().addDays(30))

        self.project_phase_input = QComboBox()
        self.project_phase_input.addItems(["Réalisation", "Conception", "Finition"])

        self.bip_option = QComboBox()
        self.bip_option.addItems(["Automatique", "Manuel"])

        self.cost_type_input = QComboBox()
        self.cost_type_input.addItems(["Monétaire (HT)", "Ressources"])

        self.cost_category_input = QComboBox()
        self.cost_category_input.addItems(["Engagé", "Dépensé"])

        base_period_layout.addWidget(QLabel("Nom du Projet:"))
        base_period_layout.addWidget(self.project_name_input)
        base_period_layout.addWidget(QLabel("Description:"))
        base_period_layout.addWidget(self.project_desc_input)
        base_period_layout.addWidget(QLabel("Phase du Projet:"))
        base_period_layout.addWidget(self.project_phase_input)
        base_period_layout.addWidget(QLabel("Date de Début:"))
        base_period_layout.addWidget(self.start_date_input)
        base_period_layout.addWidget(QLabel("Date de Fin:"))
        base_period_layout.addWidget(self.end_date_input)
        base_period_layout.addWidget(QLabel("Mode BiP:"))
        base_period_layout.addWidget(self.bip_option)
        base_period_layout.addWidget(QLabel("Type de Coût:"))
        base_period_layout.addWidget(self.cost_type_input)
        base_period_layout.addWidget(QLabel("Catégorie de Coût:"))
        base_period_layout.addWidget(self.cost_category_input)

        base_period_tab.setLayout(base_period_layout)
        self.tab_widget.addTab(base_period_tab, "Période de Base")

        # Onglet Périodes Supplémentaires
        additional_period_tab = QWidget()
        additional_period_layout = QVBoxLayout()

        self.bip_date_input = QDateEdit()
        self.bip_date_input.setDate(QDate.currentDate().addYears(1))

        additional_period_layout.addWidget(QLabel("Nouvelle Date BiP:"))
        additional_period_layout.addWidget(self.bip_date_input)

        self.additional_period_button = QPushButton("Ajouter Période Supplémentaire", self)
        self.additional_period_button.clicked.connect(self.add_additional_period)
        additional_period_layout.addWidget(self.additional_period_button, alignment=Qt.AlignCenter)

        additional_period_tab.setLayout(additional_period_layout)
        self.tab_widget.addTab(additional_period_tab, "Périodes Supplémentaires")

        layout.addWidget(self.tab_widget)

        # Navigation entre les onglets
        nav_layout = QHBoxLayout()

        prev_button = QPushButton("← Précédent", self)
        prev_button.clicked.connect(self.go_previous)
        next_button = QPushButton("Suivant →", self)
        next_button.clicked.connect(self.go_next)

        nav_layout.addWidget(prev_button)
        nav_layout.addWidget(next_button)
        layout.addLayout(nav_layout)

        # Bouton de création de projet
        create_button = QPushButton("Créer Nouveau Projet", self)
        create_button.clicked.connect(self.create_project)
        layout.addWidget(create_button, alignment=Qt.AlignCenter)

        self.setLayout(layout)

    def create_project(self):
        if self.main_window.role_manager.is_admin():
            QMessageBox.information(self, "Projet", "Projet créé avec succès !")
            self.main_window.open_project_analysis_screen()
        else:
            QMessageBox.warning(self, "Accès refusé", "Seul l'administrateur peut créer un projet.")

    def add_additional_period(self):
        QMessageBox.information(self, "Période Supplémentaire",
                                f"Nouvelle période ajoutée: {self.bip_date_input.date().toString(Qt.ISODate)}")

    def go_previous(self):
        self.main_window.return_to_main_menu()

    def go_next(self):
        current_index = self.tab_widget.currentIndex()
        if current_index < self.tab_widget.count() - 1:
            self.tab_widget.setCurrentIndex(current_index + 1)

class ProjectAnalysisScreen(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Barre d'outils
        toolbar_layout = QHBoxLayout()

        # Bouton Annuler
        undo_button = QPushButton("Annuler")
        undo_button.clicked.connect(self.undo_action)
        toolbar_layout.addWidget(undo_button)

        # Bouton Refaire
        redo_button = QPushButton("Refaire")
        redo_button.clicked.connect(self.redo_action)
        toolbar_layout.addWidget(redo_button)

        # Bouton Gras
        bold_button = QPushButton("Gras")
        bold_button.clicked.connect(self.bold_text)
        toolbar_layout.addWidget(bold_button)

        # Bouton Italique
        italic_button = QPushButton("Italique")
        italic_button.clicked.connect(self.italic_text)
        toolbar_layout.addWidget(italic_button)

        # Bouton Somme
        sum_button = QPushButton("Somme")
        sum_button.clicked.connect(self.calculate_sum)
        toolbar_layout.addWidget(sum_button)

        # Bouton Imprimer
        print_button = QPushButton("Imprimer")
        print_button.clicked.connect(self.print_table)
        toolbar_layout.addWidget(print_button)

        # Bouton Exporter en Word
        export_word_button = QPushButton("Exporter en Word")
        export_word_button.clicked.connect(self.export_table_to_word)
        toolbar_layout.addWidget(export_word_button)

        # Bouton Précédent
        back_button = QPushButton("Menu Principal")
        back_button.clicked.connect(self.main_window.return_to_main_menu)
        toolbar_layout.addWidget(back_button)

        # Spacer pour aligner les boutons à gauche
        toolbar_layout.addStretch()

        layout.addLayout(toolbar_layout)

        # Tableau
        self.table_widget = QTableWidget(5, 12)
        self.table_widget.setHorizontalHeaderLabels([
            "Nom", "Coût Estimé", "Coût Réel", "Reste à Faire", "Budget Initial",
            "Budget à Date", "Dépenses Réelles", "Valeur Acquise", "Écart", "RàF",
            "CPI", "SPI"
        ])
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table_widget)

        # Boutons supplémentaires
        button_layout = QHBoxLayout()

        calculate_button = QPushButton("Calculer les Coûts")
        calculate_button.clicked.connect(self.calculate_costs)
        button_layout.addWidget(calculate_button)

        generate_graph_button = QPushButton("Générer Graphique")
        generate_graph_button.clicked.connect(self.generate_s_curve)
        button_layout.addWidget(generate_graph_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def undo_action(self):
        QMessageBox.information(self, "Annuler", "Action Annuler (non implémentée)")

    def redo_action(self):
        QMessageBox.information(self, "Refaire", "Action Refaire (non implémentée)")

    def bold_text(self):
        selected_items = self.table_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Aucun Sélection", "Veuillez sélectionner des cellules à mettre en gras.")
            return
        for item in selected_items:
            font = item.font()
            font.setBold(True)
            item.setFont(font)

    def italic_text(self):
        selected_items = self.table_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Aucun Sélection", "Veuillez sélectionner des cellules à mettre en italique.")
            return
        for item in selected_items:
            font = item.font()
            font.setItalic(True)
            item.setFont(font)

    def calculate_sum(self):
        selected_items = self.table_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Aucun Sélection", "Veuillez sélectionner des cellules pour calculer la somme.")
            return
        total = 0
        for item in selected_items:
            try:
                total += float(item.text())
            except ValueError:
                pass
        QMessageBox.information(self, "Somme", f"Somme des cellules sélectionnées: {total}")

    def print_table(self):
        path, _ = QFileDialog.getSaveFileName(self, "Enregistrer en Excel", "", "Fichiers Excel (*.xlsx)")
        if path:
            self.export_to_excel(path)

    def export_to_excel(self, path):
        try:
            row_count = self.table_widget.rowCount()
            col_count = self.table_widget.columnCount()
            data = []

            headers = [self.table_widget.horizontalHeaderItem(col).text() for col in range(col_count)]
            data.append(headers)

            for row in range(row_count):
                row_data = []
                for col in range(col_count):
                    item = self.table_widget.item(row, col)
                    row_data.append(item.text() if item else "")
                data.append(row_data)

            df = pd.DataFrame(data[1:], columns=data[0])
            df.to_excel(path, index=False)
            QMessageBox.information(self, "Export Réussi", f"Tableau exporté avec succès vers {path}")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de l'exportation: {e}")

    def export_table_to_word(self):
        path, _ = QFileDialog.getSaveFileName(self, "Enregistrer en Word", "", "Fichiers Word (*.docx)")
        if path:
            try:
                doc = Document()

                row_count = self.table_widget.rowCount()
                col_count = self.table_widget.columnCount()

                table = doc.add_table(rows=row_count + 1, cols=col_count)

                for col in range(col_count):
                    table.cell(0, col).text = self.table_widget.horizontalHeaderItem(col).text()

                for row in range(1, row_count + 1):
                    for col in range(col_count):
                        item = self.table_widget.item(row - 1, col)
                        table.cell(row, col).text = item.text() if item else ""

                doc.save(path)
                QMessageBox.information(self, "Export Réussi", f"Tableau exporté avec succès vers {path}")
            except Exception as e:
                QMessageBox.critical(self, "Erreur", f"Erreur lors de l'exportation: {e}")

    def calculate_costs(self):
        try:
            row_count = self.table_widget.rowCount()
            for row in range(row_count):
                estimated_cost = self.get_table_value(row, 1)
                actual_cost = self.get_table_value(row, 2)
                remaining_work = estimated_cost - actual_cost if estimated_cost is not None and actual_cost is not None else 0
                self.table_widget.setItem(row, 3, QTableWidgetItem(str(remaining_work)))

                bi = self.get_table_value(row, 4)
                bad = self.get_table_value(row, 5)
                raf = self.get_table_value(row, 9)

                cpi = bi / actual_cost if actual_cost and actual_cost != 0 else 0
                spi = bad / remaining_work if remaining_work and remaining_work != 0 else 0

                self.table_widget.setItem(row, 10, QTableWidgetItem(f"{cpi:.2f}"))
                self.table_widget.setItem(row, 11, QTableWidgetItem(f"{spi:.2f}"))

            QMessageBox.information(self, "Résultat des Coûts", "Les coûts ont été calculés avec succès !")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors du calcul des coûts: {e}")

    def generate_s_curve(self):
        try:
            temps = np.linspace(0, 12, 120)
            cbtp = 100 * (1 - np.exp(-0.4 * temps))
            crte = 90 * (1 - np.exp(-0.35 * temps))
            cbte = 85 * (1 - np.exp(-0.33 * temps))
            cprf = 100 - cbtp

            plt.figure(figsize=(12, 8))
            plt.plot(temps, cbtp, label="CBTP (Budget à Date)", color='blue', linestyle='--')
            plt.plot(temps, crte, label="CRTE (Dépenses Réelles)", color='red', linestyle='-.')
            plt.plot(temps, cbte, label="CBTE (Valeur Acquise)", color='green', linestyle=':')
            plt.plot(temps, cprf, label="CPRF (Coût Restant)", color='purple', linestyle='-')

            cbtp_black = 100 * (1 - np.exp(-0.3 * temps))
            plt.plot(temps, cbtp_black, label="Montant total du BàD à F prévu à To", color='black')

            plt.hlines(y=crte[-1] - cbtp_black[-1], xmin=temps[-10], xmax=temps[-1], colors='green', label="Écart final")

            plt.vlines(x=temps[-20], ymin=0, ymax=cbtp[-1], colors='orange', label="Retard (périodes supplémentaires)")

            plt.text(1, 10, "T0", fontsize=12, verticalalignment='bottom', horizontalalignment='right')
            plt.text(11, 100, "F (prévu à To)", fontsize=12, verticalalignment='bottom', horizontalalignment='left')

            plt.title("Graphique des Courbes en S avec Retard et Écart Final")
            plt.xlabel("Temps")
            plt.ylabel("Coût")
            plt.legend()
            plt.grid(True)
            plt.show()

        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de la génération des courbes en S: {e}")

    def get_table_value(self, row, column):
        item = self.table_widget.item(row, column)
        if item and item.text():
            try:
                return float(item.text())
            except ValueError:
                return 0.0
        return 0.0


if __name__ == "__main__":
    app = QApplication(sys.argv)
    apply_stylesheet(app, theme='light_teal.xml')

    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
