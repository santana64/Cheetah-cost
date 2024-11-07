import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import matplotlib.pyplot as plt
import numpy as np
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel, QTabWidget,
    QLineEdit, QComboBox, QDateEdit, QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
    QHBoxLayout, QToolButton, QMenu, QAction, QFileDialog, QSlider, QMenuBar
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QDate
from qt_material import apply_stylesheet
import mplcursors
import pandas as pd
from docx import Document

# ----- Configuration SMTP pour notifications par e-mail -----
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_USER = "votre_email@gmail.com"
EMAIL_PASSWORD = "votre_mot_de_passe"
EMAIL_RECEIVER = "destinataire@gmail.com"

# ----- Classe RoleManager -----
class RoleManager:
    def __init__(self):
        self.user_role = "Admin"  # Rôle par défaut : Admin

    def is_admin(self):
        return self.user_role == "Admin"

# ----- Classe MainWindow -----
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Cheetah Cost - Gestion des Coûts")
        self.setGeometry(100, 100, 1200, 800)
        self.role_manager = RoleManager()
        self.init_ui()

    def init_ui(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu("Fichier")
        menu_principal_action = QAction("Menu Principal", self)
        menu_principal_action.triggered.connect(self.return_to_main_menu)
        file_menu.addAction(menu_principal_action)

        export_word_action = QAction("Exporter en Word", self)
        export_word_action.triggered.connect(lambda: self.project_analysis_screen.export_table_to_word())
        file_menu.addAction(export_word_action)

        # Options d'enregistrement
        save_action = QAction("Enregistrer", self)
        save_action.triggered.connect(self.save_table)
        file_menu.addAction(save_action)

        save_as_action = QAction("Enregistrer sous", self)
        save_as_action.triggered.connect(self.save_table_as)
        file_menu.addAction(save_as_action)

        open_action = QAction("Ouvrir", self)
        open_action.triggered.connect(self.open_table)
        file_menu.addAction(open_action)

        edit_menu = menubar.addMenu("Édition")
        add_row_action = QAction("Ajouter Ligne", self)
        add_row_action.triggered.connect(lambda: self.project_analysis_screen.add_row())
        edit_menu.addAction(add_row_action)

        remove_row_action = QAction("Supprimer Ligne", self)
        remove_row_action.triggered.connect(lambda: self.project_analysis_screen.remove_row())
        edit_menu.addAction(remove_row_action)

        view_menu = menubar.addMenu("Affichage")
        new_page_action = QAction("Nouvelle Page", self)
        new_page_action.triggered.connect(lambda: self.project_analysis_screen.create_new_table_page())
        view_menu.addAction(new_page_action)

        tools_menu = menubar.addMenu("Outils")
        bip_button_action = QAction("Ajouter Période BiP", self)
        bip_button_action.triggered.connect(self.open_project_creation)
        tools_menu.addAction(bip_button_action)

        help_menu = menubar.addMenu("Aide")
        help_action = QAction("Aide", self)
        help_menu.addAction(help_action)

        self.welcome_screen = QWidget()
        self.setCentralWidget(self.welcome_screen)
        layout = QVBoxLayout()

        label = QLabel("Bienvenue sur Cheetah Cost")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 32px; font-weight: bold; color: #3E4A59; padding: 20px;")
        layout.addWidget(label)

        start_button = QPushButton("Commencer", self)
        start_button.clicked.connect(self.open_project_creation)
        start_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 18px;
                font-weight: bold;
                padding: 15px;
                border-radius: 12px;
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

    def save_table(self):
        if self.project_analysis_screen.current_file_path:
            self.project_analysis_screen.save_to_csv(self.project_analysis_screen.current_file_path)
        else:
            self.save_table_as()

    def save_table_as(self):
        path, _ = QFileDialog.getSaveFileName(self, "Enregistrer le tableau", "", "CSV Files (*.csv)")
        if path:
            self.project_analysis_screen.save_to_csv(path)

    def open_table(self):
        path, _ = QFileDialog.getOpenFileName(self, "Ouvrir le tableau", "", "CSV Files (*.csv)")
        if path:
            self.project_analysis_screen.load_from_csv(path)

# ----- Fonction pour envoyer des notifications par e-mail -----
def envoyer_email_notification(sujet, corps):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_USER
    msg['To'] = EMAIL_RECEIVER
    msg['Subject'] = sujet
    msg.attach(MIMEText(corps, 'plain'))

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASSWORD)
        server.sendmail(EMAIL_USER, EMAIL_RECEIVER, msg.as_string())
        server.quit()
        print("Email envoyé avec succès!")
    except Exception as e:
        print(f"Erreur lors de l'envoi de l'email: {e}")

# ----- Classe ProjectCreationScreen -----
class ProjectCreationScreen(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.bip_dates = []
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        self.tab_widget = QTabWidget()

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

        additional_period_tab = QWidget()
        additional_period_layout = QVBoxLayout()

        self.bip_date_input = QDateEdit()
        self.bip_date_input.setDate(QDate.currentDate().addYears(1))

        additional_period_layout.addWidget(QLabel("Nouvelle Date BiP:"))
        additional_period_layout.addWidget(self.bip_date_input)

        self.additional_period_button = QPushButton("Ajouter Période Supplémentaire", self)
        self.additional_period_button.clicked.connect(self.add_additional_period)
        additional_period_layout.addWidget(self.additional_period_button, alignment=Qt.AlignCenter)

        self.bip_table = QTableWidget(0, 1)
        self.bip_table.setHorizontalHeaderLabels(["Dates BiP"])
        self.bip_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        additional_period_layout.addWidget(self.bip_table)

        additional_period_tab.setLayout(additional_period_layout)
        self.tab_widget.addTab(additional_period_tab, "Périodes Supplémentaires")

        layout.addWidget(self.tab_widget)
        nav_layout = QHBoxLayout()

        prev_button = QPushButton("← Précédent", self)
        prev_button.clicked.connect(self.go_previous)
        next_button = QPushButton("Suivant →", self)
        next_button.clicked.connect(self.go_next)

        nav_layout.addWidget(prev_button)
        nav_layout.addWidget(next_button)
        layout.addLayout(nav_layout)

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
        bip_date = self.bip_date_input.date().toString(Qt.ISODate)
        if bip_date in self.bip_dates:
            QMessageBox.warning(self, "Erreur", "Une période avec cette date existe déjà.")
        else:
            self.bip_dates.append(bip_date)
            row_position = self.bip_table.rowCount()
            self.bip_table.insertRow(row_position)
            self.bip_table.setItem(row_position, 0, QTableWidgetItem(bip_date))
            envoyer_email_notification("Notification Période BiP", f"Nouvelle période BiP ajoutée : {bip_date}")
            QMessageBox.information(self, "Période Supplémentaire", f"Nouvelle période ajoutée: {bip_date}")

    def go_previous(self):
        self.main_window.return_to_main_menu()

    def go_next(self):
        current_index = self.tab_widget.currentIndex()
        if current_index < self.tab_widget.count() - 1:
            self.tab_widget.setCurrentIndex(current_index + 1)

# ----- Classe ProjectAnalysisScreen -----
class ProjectAnalysisScreen(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.current_file_path = None  # Pour enregistrer le chemin du fichier courant
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Tableau avec colonnes élargies
        self.table_widget = QTableWidget(5, 12)  # Initialement 5 lignes, maximum 20
        self.table_widget.setHorizontalHeaderLabels([
            "Nom", "Coût Estimé", "Coût Réel", "Reste à Faire", "Budget Initial",
            "Budget à Date", "Dépenses Réelles", "Valeur Acquise", "Écart", "RàF",
            "CPI", "SPI"
        ])
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        layout.addWidget(self.table_widget)

        # Boutons supplémentaires
        button_layout = QHBoxLayout()

        calculate_button = QPushButton("Calculer les Coûts", self)
        calculate_button.clicked.connect(self.calculate_costs)
        button_layout.addWidget(calculate_button)

        generate_graph_button = QPushButton("Générer Graphique", self)
        generate_graph_button.clicked.connect(self.generate_s_curve)
        button_layout.addWidget(generate_graph_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def add_row(self):
        if self.table_widget.rowCount() < 20:
            self.table_widget.insertRow(self.table_widget.rowCount())
        else:
            QMessageBox.warning(self, "Limite atteinte", "Vous ne pouvez pas ajouter plus de 20 lignes.")

    def remove_row(self):
        if self.table_widget.rowCount() > 1:
            self.table_widget.removeRow(self.table_widget.rowCount() - 1)
        else:
            QMessageBox.warning(self, "Limite atteinte", "Il doit rester au moins une ligne.")

    def calculate_costs(self):
        try:
            row_count = self.table_widget.rowCount()
            for row in range(row_count):
                estimated_cost = self.get_table_value(row, 1)
                actual_cost = self.get_table_value(row, 2)
                remaining_work = estimated_cost - actual_cost if estimated_cost and actual_cost else 0
                self.table_widget.setItem(row, 3, QTableWidgetItem(str(remaining_work)))

                bi = self.get_table_value(row, 4)
                raf = remaining_work  # Calcul du reste à faire
                self.table_widget.setItem(row, 9, QTableWidgetItem(str(raf)))

                cost_variance = bi - actual_cost
                self.table_widget.setItem(row, 8, QTableWidgetItem(str(cost_variance)))

                # Calculs des indices CPI et SPI
                cpi = bi / actual_cost if actual_cost else 0
                spi = raf / remaining_work if remaining_work else 0
                self.table_widget.setItem(row, 10, QTableWidgetItem(f"{cpi:.2f}"))
                self.table_widget.setItem(row, 11, QTableWidgetItem(f"{spi:.2f}"))

                if actual_cost > estimated_cost:
                    envoyer_email_notification("Dépassement de Budget", f"Le coût réel ({actual_cost}) dépasse le budget estimé ({estimated_cost})")

            QMessageBox.information(self, "Résultat des Coûts", "Les coûts ont été calculés avec succès !")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors du calcul des coûts: {e}")

    def generate_s_curve(self):
        try:
            row_count = self.table_widget.rowCount()
            b2p, budget, cp, depenses, val_acq = [], [], [], [], []

            for row in range(row_count):
                b2p.append(row)
                budget.append(self.get_table_value(row, 4))
                cp.append(self.get_table_value(row, 5))
                depenses.append(self.get_table_value(row, 6))
                val_acq.append(self.get_table_value(row, 7))

            plt.figure(figsize=(12, 8))
            plt.plot(b2p, budget, label="Budget", color='black', linestyle='-')
            plt.plot(b2p, cp, label="CP", color='green', linestyle='-')
            plt.plot(b2p, depenses, label="Dépenses", color='red', linestyle='-')
            plt.plot(b2p, val_acq, label="Valeur Acquise", color='blue', linestyle='-')
            plt.title("Méthode FGF de Coûtenance - Courbes")
            plt.xlabel("B2P n°")
            plt.ylabel("Valeur")
            plt.legend()
            plt.grid(True)

            cursor = mplcursors.cursor(hover=True)
            cursor.connect("add", lambda sel: sel.annotation.set_text(f"Valeur: {sel.target[1]:.2f}"))

            plt.show()
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de la génération des courbes en FGF: {e}")

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
                QMessageBox.information(self, "Export Réussie", f"Tableau exporté avec succès vers {path}")
            except Exception as e:
                QMessageBox.critical(self, "Erreur", f"Erreur lors de l'exportation: {e}")

    def save_to_csv(self, path):
        try:
            row_count = self.table_widget.rowCount()
            col_count = self.table_widget.columnCount()
            data = []

            for row in range(row_count):
                row_data = []
                for col in range(col_count):
                    item = self.table_widget.item(row, col)
                    row_data.append(item.text() if item else "")
                data.append(row_data)

            df = pd.DataFrame(data, columns=[self.table_widget.horizontalHeaderItem(i).text() for i in range(col_count)])
            df.to_csv(path, index=False)
            self.current_file_path = path
            QMessageBox.information(self, "Enregistrement réussi", f"Tableau enregistré sous {path}")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de l'enregistrement: {e}")

    def load_from_csv(self, path):
        try:
            df = pd.read_csv(path).fillna("")  # Remplace les valeurs NaN par des chaînes vides
            self.table_widget.setRowCount(df.shape[0])
            self.table_widget.setColumnCount(df.shape[1])
            self.table_widget.setHorizontalHeaderLabels(df.columns.tolist())

            for row in range(df.shape[0]):
                for col in range(df.shape[1]):
                    item = QTableWidgetItem(str(df.iat[row, col]))
                    self.table_widget.setItem(row, col, item)

            self.current_file_path = path
            QMessageBox.information(self, "Ouverture réussie", f"Tableau chargé depuis {path}")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de l'ouverture: {e}")

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
