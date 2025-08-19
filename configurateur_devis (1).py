import sys
import os
sys.path.append(r"C:\cadwork\libs")

import cadwork
import geometry_controller as gc
import attribute_controller as ac
import element_controller as ec
import material_controller as mc
import utility_controller as uc
from datetime import datetime
from collections import defaultdict
import subprocess
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog
import threading
import json

class ConfigurateurDevisInterface:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("⚙️ CONFIGURATEUR DE DEVIS - Cadwork v2.0")
        self.root.geometry("1400x900")
        self.root.configure(bg='#f0f0f0')
        
        # Variables principales
        self.element_ids = []
        self.materiaux_detectes = {}  # {material_name: {'count': int, 'exemple_id': int, 'prix_cadwork': float}}
        self.config_vars = {}  # Variables Tkinter pour chaque matériau
        self.should_stop = False
        self.config_validee = False
        
        # Constantes de prix existantes (pour référence)
        self.prix_traitement = {"CL2": 25.5, "CL3": 215, "CL2 INCOLORE": 50}
        self.prix_faconnage = {
            "T1_V": 224.4, "T16_L": 2.34, "T2_V": 243.1, "T3_V": 320, "T4_V": 263.5
        }
        self.prix_prestation = {
            "MO1_S": 17, "MO2_S": 28.9, "MO3_S": 22.95, "MO4_S": 19.55,
            "MO5_V": 219.3, "MO6_V": 85, "MO7_V": 119, "MO8_V": 35
        }
        
        self.setup_ui()
        
    def setup_ui(self):
        """Configure l'interface utilisateur"""
        # Style
        style = ttk.Style()
        style.theme_use('clam')
        
        # Titre principal
        title_frame = tk.Frame(self.root, bg='#2c3e50', height=80)
        title_frame.pack(fill='x', padx=0, pady=0)
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(title_frame, text="⚙️ CONFIGURATEUR DE DEVIS", 
                              font=('Arial', 20, 'bold'), fg='white', bg='#2c3e50')
        title_label.pack(expand=True)
        
        # Frame principal
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Zone d'information projet
        self.create_project_info_frame(main_frame)
        
        # Boutons d'action principaux
        self.create_action_buttons_frame(main_frame)
        
        # Zone de configuration des matériaux (avec scroll)
        self.create_material_config_frame(main_frame)
        
        # Zone de log
        self.create_log_frame(main_frame)
        
        # Bouton fermer
        self.create_close_button_frame(main_frame)
        
        # Gestionnaire de fermeture de fenêtre
        self.root.protocol("WM_DELETE_WINDOW", self.close_application)
        
        # Initialisation
        self.update_project_info()
        
    def create_project_info_frame(self, parent):
        """Crée la zone d'information du projet"""
        info_frame = tk.LabelFrame(parent, text="📊 Informations du projet", 
                                 font=('Arial', 12, 'bold'), bg='#f0f0f0', fg='#2c3e50')
        info_frame.pack(fill='x', pady=(0, 10))
        
        self.project_info_label = tk.Label(info_frame, text="En attente d'analyse...", 
                                         font=('Arial', 11), bg='#f0f0f0', fg='#3498db')
        self.project_info_label.pack(pady=10, padx=10)
        
    def create_action_buttons_frame(self, parent):
        """Crée les boutons d'action principaux"""
        action_frame = tk.Frame(parent, bg='#f0f0f0')
        action_frame.pack(fill='x', pady=(0, 10))
        
        self.analyze_button = tk.Button(action_frame, text="🔍 ANALYSER LES MATÉRIAUX", 
                                      command=self.analyze_materials, font=('Arial', 12, 'bold'),
                                      bg='#3498db', fg='white', height=2, cursor='hand2')
        self.analyze_button.pack(side='left', padx=(0, 10))
        
        self.validate_button = tk.Button(action_frame, text="✅ VALIDER CONFIGURATION", 
                                       command=self.validate_config, font=('Arial', 12, 'bold'),
                                       bg='#27ae60', fg='white', height=2, cursor='hand2',
                                       state='disabled')
        self.validate_button.pack(side='left', padx=(0, 10))
        
        self.optimize_button = tk.Button(action_frame, text="🚀 CONFIGURER OPTIMISATION", 
                                       command=self.launch_optimization, font=('Arial', 12, 'bold'),
                                       bg='#e67e22', fg='white', height=2, cursor='hand2',
                                       state='disabled')
        self.optimize_button.pack(side='left', padx=(0, 10))
        
    def create_material_config_frame(self, parent):
        """Crée la zone de configuration des matériaux avec scroll"""
        config_frame = tk.LabelFrame(parent, text="⚙️ Configuration des matériaux", 
                                   font=('Arial', 12, 'bold'), bg='#f0f0f0', fg='#2c3e50')
        config_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Instructions avec explication de la logique optimisation/méthode - CORRIGÉ
        instructions = tk.Label(config_frame, 
                              text="Configurez chaque matériau selon vos besoins. Les prix seront automatiquement récupérés depuis Cadwork.\n"
                                   "💡 Logique simplifiée : Matière/Façonnage/Traitement/Prestations calculés automatiquement selon les codes présents.\n"
                                   "💡 Chutes = Automatiques si optimisation, sinon 0%. Paramètres détaillés d'optimisation à l'étape suivante.\n"
                                   "🎯 MÉTHODES COHÉRENTES : m³→Volume, m²→Surface, ml→Longueur, Kg→Poids, U→Quantité. Manuel si optimisé.\n"
                                   "📐 Nouvelles surfaces : Face référence=API, Surface réelle=Volume physique/épaisseur.",
                              font=('Arial', 10), bg='#f0f0f0', fg='#7f8c8d', justify='left')
        instructions.pack(pady=(10, 5), padx=10, anchor='w')
        
        # Frame avec scrollbar pour la configuration
        canvas_frame = tk.Frame(config_frame, bg='#f0f0f0')
        canvas_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.canvas = tk.Canvas(canvas_frame, bg='white', highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg='white')
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mousewheel to canvas
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        
    def create_log_frame(self, parent):
        """Crée la zone de log"""
        log_frame = tk.LabelFrame(parent, text="📋 Journal", font=('Arial', 12, 'bold'), 
                                bg='#f0f0f0', fg='#2c3e50')
        log_frame.pack(fill='x', pady=(10, 0))
        
        self.text_area = scrolledtext.ScrolledText(log_frame, height=8, width=80, 
                                                 font=('Courier', 9), bg='#ffffff', 
                                                 fg='#2c3e50', wrap=tk.WORD)
        self.text_area.pack(fill='x', padx=10, pady=10)
        
    def create_close_button_frame(self, parent):
        """Crée le bouton fermer"""
        close_frame = tk.Frame(parent, bg='#f0f0f0')
        close_frame.pack(fill='x', pady=(10, 0))
        
        self.close_button = tk.Button(close_frame, text="🚪 FERMER", 
                                    command=self.close_application, font=('Arial', 12, 'bold'),
                                    bg='#95a5a6', fg='white', height=2, cursor='hand2')
        self.close_button.pack(side='right')
        
    def _on_mousewheel(self, event):
        """Gère la molette de la souris pour le scroll"""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    def log(self, message):
        """Ajoute un message dans la zone de texte"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.text_area.config(state='normal')
        self.text_area.insert(tk.END, f"[{timestamp}] {message}\n")
        self.text_area.config(state='disabled')
        self.text_area.see(tk.END)
        self.root.update()
        
    def update_project_info(self):
        """Met à jour les informations du projet"""
        try:
            project_num = uc.get_project_number() or "N/A"
            project_attr = uc.get_project_user_attribute(1) or "N/A"
            elements_count = len(self.element_ids)
            materials_count = len(self.materiaux_detectes)
            
            info_text = f"📊 Projet: {project_num} | Attribut: {project_attr} | Éléments: {elements_count} | Matériaux: {materials_count}"
            self.project_info_label.config(text=info_text)
            
        except Exception as e:
            self.project_info_label.config(text=f"Erreur lecture projet: {e}")
            
    def safe_float(self, val, default=0.0):
        """Conversion sécurisée en float"""
        try:
            return float(str(val).replace(',', '.').replace('%', ''))
        except (ValueError, TypeError):
            return default
            
    def get_material_unit_auto(self, material_name):
        """Détecte automatiquement l'unité d'un matériau avec logique affinée"""
        try:
            mat_id = mc.get_material_id(material_name)
            if hasattr(mc, 'get_unit'):
                unit = mc.get_unit(mat_id)
                if unit:
                    return unit.lower()
        except Exception:
            pass
        
        # Heuristique basée sur le nom avec plus de précision
        up = (material_name or '').upper()
        
        # Quincaillerie et pièces unitaires
        if any(k in up for k in ['VIS', 'BOULON', 'CLOU', 'CHEVILLE', 'EQUERRE', 'SABOT', 'CONNECTEUR', 'ASSEMBLAGE']):
            return 'U'
        
        # Matériaux au poids (métaux, acier)
        if any(k in up for k in ['ACIER', 'METAL', 'INOX', 'PLOMB', 'ZINC', 'ALUMINIUM', 'LAITON']):
            return 'Kg'
            
        # Panneaux et surfaces (priorité sur linéaires pour _L dans panneaux)
        if any(k in up for k in ['OSB', 'CONTREPLAQUE', 'PANNEAU', 'PLAQUE', 'ISOLANT', 'MEMBRANE', 'BARDAGE']):
            return 'm²'
            
        # Linéaires (après panneaux pour éviter confusion)
        if any(k in up for k in ['POUTRE', 'PROFILE', 'CHEVRON', 'TASSEAU', 'SJ-', '_L']) and not any(k in up for k in ['PANNEAU', 'PLAQUE']):
            return 'ml'
            
        # Volume par défaut (bois, etc.)
        return 'm³'
        
    def get_cadwork_material_price(self, material_name):
        """Récupère le prix depuis les matériaux Cadwork"""
        try:
            mat_id = mc.get_material_id(material_name)
            prix = self.safe_float(mc.get_price(mat_id))
            return prix
        except Exception as e:
            self.log(f"⚠️ Impossible de récupérer le prix pour {material_name}: {e}")
            return 0.0
            
    def analyze_materials(self):
        """Analyse les éléments et détecte tous les matériaux"""
        try:
            self.log("🔍 Récupération des éléments sélectionnés...")
            self.element_ids = ec.get_active_identifiable_element_ids()
            
            if not self.element_ids:
                messagebox.showwarning("Aucun élément", 
                                     "Veuillez sélectionner des éléments dans Cadwork avant l'analyse.")
                return
                
            self.log(f"📊 Analyse de {len(self.element_ids)} élément(s)")
            
            # Détection des matériaux
            materiaux_stats = defaultdict(lambda: {'count': 0, 'exemple_id': None, 'prix_cadwork': 0.0})
            
            for eid in self.element_ids:
                try:
                    mat_name = ac.get_element_material_name(eid).strip()
                    if mat_name:
                        materiaux_stats[mat_name]['count'] += 1
                        if materiaux_stats[mat_name]['exemple_id'] is None:
                            materiaux_stats[mat_name]['exemple_id'] = eid
                            # Récupération du prix Cadwork
                            materiaux_stats[mat_name]['prix_cadwork'] = self.get_cadwork_material_price(mat_name)
                except Exception as e:
                    self.log(f"⚠️ Erreur élément {eid}: {e}")
                    continue
            
            self.materiaux_detectes = dict(materiaux_stats)
            self.log(f"✅ {len(self.materiaux_detectes)} matériau(x) détecté(s)")
            
            # Création de l'interface de configuration
            self.create_config_interface()
            
            # Sauvegarde des prix dans les attributs
            self.save_material_prices_to_attributes()
            
            # Mise à jour des informations
            self.update_project_info()
            
            # Activation des boutons
            self.validate_button.config(state='normal')
            
        except Exception as e:
            self.log(f"❌ Erreur lors de l'analyse : {e}")
            messagebox.showerror("Erreur", f"Erreur lors de l'analyse :\n{e}")
    
    def save_material_prices_to_attributes(self):
        """Sauvegarde les prix matière dans l'attribut 19 de tous les éléments"""
        try:
            self.log("💾 Sauvegarde des prix matière dans les attributs Cadwork...")
            
            for eid in self.element_ids:
                try:
                    mat_name = ac.get_element_material_name(eid).strip()
                    if mat_name in self.materiaux_detectes:
                        prix_cadwork = self.materiaux_detectes[mat_name]['prix_cadwork']
                        ac.set_user_attribute([eid], 19, str(prix_cadwork))
                except Exception as e:
                    self.log(f"⚠️ Erreur sauvegarde prix élément {eid}: {e}")
                    
            self.log("✅ Prix matière sauvegardés dans attribut 19")
            
        except Exception as e:
            self.log(f"❌ Erreur sauvegarde prix : {e}")
    
    def create_config_interface(self):
        """Crée l'interface de configuration pour chaque matériau"""
        # Nettoyer le frame existant
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        self.config_vars = {}
        
        # Configuration des colonnes avec largeurs fixes (ajustée pour méthodes plus longues)
        self.col_widths = [240, 50, 80, 90, 60, 180, 80, 100]  # Largeurs en pixels
        
        # Headers du tableau
        header_frame = tk.Frame(self.scrollable_frame, bg='#34495e', relief='ridge', bd=2)
        header_frame.pack(fill='x', padx=5, pady=5)
        
        # Configuration des colonnes du header
        for i, width in enumerate(self.col_widths):
            header_frame.grid_columnconfigure(i, minsize=width, weight=0)
        
        headers = ["Matériau", "Qté", "Optimiser", "Type Opti", "Unité", "Méthode calcul", "Prix €/u", "Actions"]
        
        for col, title in enumerate(headers):
            lbl = tk.Label(header_frame, text=title, font=('Arial', 10, 'bold'), 
                          bg='#34495e', fg='white')
            lbl.grid(row=0, column=col, padx=2, pady=5, sticky='ew')
        
        # Configuration pour chaque matériau
        row_num = 1
        for mat_name, stats in self.materiaux_detectes.items():
            self.create_material_row(mat_name, stats, row_num)
            row_num += 1
        
        # Boutons de configuration rapide
        self.create_quick_config_buttons()
        
    def create_material_row(self, mat_name, stats, row_num):
        """Crée une ligne de configuration pour un matériau"""
        self.config_vars[mat_name] = {}
        
        # Frame pour la ligne
        row_color = '#ecf0f1' if row_num % 2 == 0 else '#ffffff'
        mat_frame = tk.Frame(self.scrollable_frame, bg=row_color, relief='ridge', bd=1)
        mat_frame.pack(fill='x', padx=5, pady=1)
        
        # Configuration des colonnes identique au header
        for i, width in enumerate(self.col_widths):
            mat_frame.grid_columnconfigure(i, minsize=width, weight=0)
        
        # Nom du matériau
        tk.Label(mat_frame, text=mat_name[:35], font=('Arial', 9), 
                bg=row_color, anchor='w').grid(row=0, column=0, padx=2, pady=2, sticky='ew')
        
        # Quantité
        tk.Label(mat_frame, text=str(stats['count']), font=('Arial', 9), 
                bg=row_color).grid(row=0, column=1, padx=2, pady=2, sticky='ew')
        
        # À optimiser - Frame pour centrage parfait
        opt_frame = tk.Frame(mat_frame, bg=row_color)
        opt_frame.grid(row=0, column=2, padx=2, pady=2, sticky='ew')
        self.config_vars[mat_name]['optimiser'] = tk.BooleanVar()
        tk.Checkbutton(opt_frame, variable=self.config_vars[mat_name]['optimiser'],
                      bg=row_color, command=lambda m=mat_name: self.on_optimiser_changed(m)).pack(expand=True)
        
        # Type optimisation
        self.config_vars[mat_name]['type_opti'] = tk.StringVar(value="variable")
        self.config_vars[mat_name]['type_opti_combo'] = ttk.Combobox(mat_frame, textvariable=self.config_vars[mat_name]['type_opti'],
                                values=["fixe", "variable"], width=8, state="disabled")
        self.config_vars[mat_name]['type_opti_combo'].grid(row=0, column=3, padx=2, pady=2, sticky='ew')
        self.config_vars[mat_name]['type_opti_combo'].bind("<<ComboboxSelected>>", lambda e, m=mat_name: self.on_config_changed(m))
        
        # Unité
        self.config_vars[mat_name]['unite'] = tk.StringVar()
        unite_combo = ttk.Combobox(mat_frame, textvariable=self.config_vars[mat_name]['unite'],
                                 values=["m³", "m²", "ml", "Kg", "U"], width=5, state="readonly")
        unite_combo.grid(row=0, column=4, padx=2, pady=2, sticky='ew')
        unite_combo.bind("<<ComboboxSelected>>", lambda e, m=mat_name: self.on_unite_changed(m))
        
        # Méthode de calcul avec noms explicites
        self.config_vars[mat_name]['methode_calcul'] = tk.StringVar(value="Manuel (L×l×h)")
        self.config_vars[mat_name]['methode_combo'] = ttk.Combobox(mat_frame, textvariable=self.config_vars[mat_name]['methode_calcul'],
                                   values=["Manuel (L×l×h)"], width=18, state="readonly")
        self.config_vars[mat_name]['methode_combo'].grid(row=0, column=5, padx=2, pady=2, sticky='ew')
        self.config_vars[mat_name]['methode_combo'].bind("<<ComboboxSelected>>", lambda e, m=mat_name: self.on_config_changed(m))
        
        # Prix unitaire
        self.config_vars[mat_name]['prix_unitaire'] = tk.StringVar(value=f"{stats['prix_cadwork']:.2f}")
        prix_entry = tk.Entry(mat_frame, textvariable=self.config_vars[mat_name]['prix_unitaire'],
                             width=8, font=('Arial', 9), justify='center')
        prix_entry.grid(row=0, column=6, padx=2, pady=2, sticky='ew')
        prix_entry.bind('<KeyRelease>', lambda e, m=mat_name: self.on_price_changed(m))
        
        # Actions
        actions_frame = tk.Frame(mat_frame, bg=row_color)
        actions_frame.grid(row=0, column=7, padx=2, pady=2, sticky='ew')
        
        tk.Button(actions_frame, text="💰", command=lambda m=mat_name: self.modify_price(m),
                 font=('Arial', 8), bg='#f39c12', fg='white', width=3).pack(side='left', padx=1)
        tk.Button(actions_frame, text="🔄", command=lambda m=mat_name: self.reset_price(m),
                 font=('Arial', 8), bg='#95a5a6', fg='white', width=3).pack(side='left', padx=1)
        
        # Configuration par défaut basée sur le nom du matériau
        self.set_default_config(mat_name)
        
    def create_quick_config_buttons(self):
        """Crée les boutons de configuration rapide"""
        separator = tk.Frame(self.scrollable_frame, height=2, bg='#bdc3c7')
        separator.pack(fill='x', padx=5, pady=10)
        
        quick_frame = tk.Frame(self.scrollable_frame, bg='white')
        quick_frame.pack(fill='x', padx=5, pady=5)
        
        tk.Label(quick_frame, text="⚡ Configuration rapide :", font=('Arial', 11, 'bold'), 
                bg='white').pack(side='left', padx=5)
        
        buttons = [
            ("🌲 Bois structure", self.config_bois_structure, '#27ae60'),
            ("📦 Panneaux", self.config_panneaux, '#f39c12'),
            ("🏠 Enveloppes", self.config_enveloppes, '#e74c3c'),
            ("📏 Linéaires", self.config_lineaires, '#9b59b6'),
            ("🔩 Quincaillerie", self.config_quincaillerie, '#34495e'),
            ("🔄 Reset", self.config_reset, '#95a5a6')
        ]
        
        for text, command, color in buttons:
            tk.Button(quick_frame, text=text, command=command,
                     bg=color, fg='white', font=('Arial', 9), cursor='hand2').pack(side='left', padx=2)
        
    def set_default_config(self, mat_name):
        """Définit une configuration par défaut basée sur le nom du matériau"""
        mat_lower = mat_name.lower()
        
        # Auto-détection de l'unité
        unite_auto = self.get_material_unit_auto(mat_name)
        self.config_vars[mat_name]['unite'].set(unite_auto)
        
        # Définition par défaut de l'optimisation selon le type de matériau
        optimise_defaut = False
        
        if any(x in mat_lower for x in ['enveloppe', 'paroi', 'isolation']):
            # Enveloppes : pas d'optimisation, surface
            optimise_defaut = False
            self.config_vars[mat_name]['unite'].set("m²")
        elif any(x in mat_lower for x in ['osb', 'panneau', 'contreplaque', 'mdf']):
            # Panneaux : pas d'optimisation, surface
            optimise_defaut = False
            self.config_vars[mat_name]['unite'].set("m²")
        elif any(x in mat_lower for x in ['_l', 'barre', 'profile']):
            # Matériaux linéaires : optimisation, longueur
            optimise_defaut = True
            self.config_vars[mat_name]['unite'].set("ml")
        elif any(x in mat_lower for x in ['vis', 'boulon', 'clou', 'cheville', 'equerre', 'sabot']):
            # Quincaillerie : pas d'optimisation, unité
            optimise_defaut = False
            self.config_vars[mat_name]['unite'].set("U")
        elif any(x in mat_lower for x in ['acier', 'metal', 'inox']):
            # Métaux : pas d'optimisation, poids
            optimise_defaut = False
            self.config_vars[mat_name]['unite'].set("Kg")
        else:
            # Bois par défaut : optimisation, volume
            optimise_defaut = True
            self.config_vars[mat_name]['unite'].set("m³")
            
        # Application de l'optimisation
        self.config_vars[mat_name]['optimiser'].set(optimise_defaut)
        
        # Mise à jour automatique des méthodes selon l'unité ET l'optimisation
        unite_finale = self.config_vars[mat_name]['unite'].get()
        methode_adaptee = self.get_methode_defaut_par_unite_et_opti(unite_finale, optimise_defaut)
        
        # Mise à jour de la liste des méthodes disponibles
        methodes_disponibles = self.get_methodes_par_unite(unite_finale)
        self.config_vars[mat_name]['methode_combo']['values'] = methodes_disponibles
        self.config_vars[mat_name]['methode_calcul'].set(methode_adaptee)
        
        # Mise à jour de l'état du combo type optimisation
        self.on_optimiser_changed(mat_name)
    
    def get_methodes_par_unite(self, unite):
        """Retourne les méthodes de calcul disponibles selon l'unité - CORRIGÉ"""
        methodes_map = {
            "m³": ["Manuel (L×l×h)", "Volume standard (brut)", "Volume liste", "Volume physique réel"],
            "m²": ["Surface (l×h)", "Surface face avant", "Surface face référence", "Surface réelle"],
            "ml": ["Manuel longueur", "Longueur liste", "Longueur physique"],
            "Kg": ["Manuel poids", "Poids matériau"],
            "U": ["Manuel quantité", "Nombre pièces"]
        }
        return methodes_map.get(unite, ["Manuel (L×l×h)"])
    
    def get_methode_defaut_par_unite_et_opti(self, unite, optimise=False):
        """Retourne la méthode par défaut selon l'unité ET le mode optimisation - CORRIGÉ"""
        if optimise:
            # Optimisé = Manuel pour maîtriser les calculs et taux de chute
            defaut_map = {
                "m³": "Manuel (L×l×h)",  # Volume contrôlé pour optimisation
                "m²": "Surface (l×h)",   # Surface contrôlée
                "ml": "Manuel longueur", # Longueur contrôlée pour optimisation  
                "Kg": "Manuel poids",    # Poids contrôlé
                "U": "Manuel quantité"   # Quantité contrôlée
            }
        else:
            # Non optimisé = API Cadwork pour précision maximale
            defaut_map = {
                "m³": "Volume standard (brut)",  # API Cadwork précise
                "m²": "Surface face référence",  # API Cadwork pour surface
                "ml": "Longueur physique",       # API Cadwork avec dépassements
                "Kg": "Poids matériau",          # API Cadwork pour poids
                "U": "Nombre pièces"             # Comptage pour unités
            }
        return defaut_map.get(unite, "Manuel (L×l×h)")
    
    def get_methode_defaut_par_unite(self, unite):
        """Méthode par défaut générique - CORRIGÉ"""
        defaut_map = {
            "m³": "Manuel (L×l×h)",  # Seule méthode qui garde ce nom car c'est un vrai volume
            "m²": "Surface (l×h)",   
            "ml": "Manuel longueur",  
            "Kg": "Manuel poids",    
            "U": "Manuel quantité"   
        }
        return defaut_map.get(unite, "Manuel (L×l×h)")
    
    def on_unite_changed(self, mat_name):
        """Callback quand l'unité change - met à jour les méthodes disponibles"""
        unite = self.config_vars[mat_name]['unite'].get()
        
        # Mettre à jour les méthodes disponibles
        methodes_disponibles = self.get_methodes_par_unite(unite)
        methode_combo = self.config_vars[mat_name]['methode_combo']
        methode_combo['values'] = methodes_disponibles
        
        # Sélectionner la méthode par défaut selon l'unité ET l'optimisation
        optimise = self.config_vars[mat_name]['optimiser'].get()
        methode_defaut = self.get_methode_defaut_par_unite_et_opti(unite, optimise)
        self.config_vars[mat_name]['methode_calcul'].set(methode_defaut)
        
        self.log(f"📏 {mat_name}: Unité → {unite}, Méthode → {methode_defaut}")
        
        # Log explicatif des méthodes selon l'unité - CORRIGÉ
        if unite == "m³":
            self.log(f"   💡 Méthodes m³: Manuel(L×l×h), Standard(brut), Liste(nomenclature), Physique(réel)")
        elif unite == "m²":
            self.log(f"   💡 Méthodes m²: Surface(l×h), Face avant(API), Face référence(API), Surface réelle")
        elif unite == "ml":
            self.log(f"   💡 Méthodes ml: Manuel longueur, Liste(nomenclature), Physique(avec dépassements)")
        elif unite == "Kg":
            self.log(f"   💡 Méthodes Kg: Manuel poids, Poids matériau(API)")
        elif unite == "U":
            self.log(f"   💡 Méthodes U: Manuel quantité, Nombre pièces(comptage)")
        
    def get_methode_code(self, methode_display):
        """Convertit le nom affiché en code technique - CORRIGÉ"""
        mapping = {
            # Volume (m³)
            "Manuel (L×l×h)": "manuel",
            "Volume standard (brut)": "volume_standard", 
            "Volume liste": "volume_liste",
            "Volume physique réel": "volume_physique_reel",
            # Surface (m²) 
            "Surface (l×h)": "surface_largeur_hauteur",
            "Surface face avant": "surface_face_avant",
            "Surface face référence": "surface_face_reference",
            "Surface réelle": "surface_reelle",
            # Longueur (ml)
            "Manuel longueur": "manuel_longueur",
            "Longueur liste": "longueur_liste",
            "Longueur physique": "longueur_physique",
            # Poids (Kg)
            "Manuel poids": "manuel_poids",
            "Poids matériau": "poids_materiau",
            # Quantité (U)
            "Manuel quantité": "manuel_quantite",
            "Nombre pièces": "nombre_pieces"
        }
        return mapping.get(methode_display, "manuel")
        
    def get_methode_display(self, methode_code):
        """Convertit le code technique en nom affiché - CORRIGÉ"""
        mapping = {
            # Volume (m³)
            "manuel": "Manuel (L×l×h)",
            "volume_standard": "Volume standard (brut)",
            "volume_liste": "Volume liste", 
            "volume_physique_reel": "Volume physique réel",
            # Surface (m²)
            "surface_largeur_hauteur": "Surface (l×h)",
            "surface_face_avant": "Surface face avant",
            "surface_face_reference": "Surface face référence",
            "surface_reelle": "Surface réelle",
            # Longueur (ml)
            "manuel_longueur": "Manuel longueur",
            "longueur_liste": "Longueur liste",
            "longueur_physique": "Longueur physique",
            # Poids (Kg)
            "manuel_poids": "Manuel poids",
            "poids_materiau": "Poids matériau",
            # Quantité (U)
            "manuel_quantite": "Manuel quantité",
            "nombre_pieces": "Nombre pièces"
        }
        return mapping.get(methode_code, "Manuel (L×l×h)")
    
    def on_optimiser_changed(self, mat_name):
        """Callback quand l'option 'Optimiser' change - active/désactive le type d'optimisation ET ajuste la méthode"""
        optimiser = self.config_vars[mat_name]['optimiser'].get()
        combo_type = self.config_vars[mat_name]['type_opti_combo']
        
        if optimiser:
            combo_type.config(state="readonly")
            # Style activé - couleur normale et remise de la valeur par défaut
            combo_type.configure(foreground='black')
            if combo_type.get() == "(désactivé)":
                combo_type.set("variable")
        else:
            combo_type.config(state="disabled")
            # Style désactivé - couleur grisée et texte explicite
            combo_type.set("(désactivé)")
            combo_type.configure(foreground='gray')
            
        # NOUVEAU : Ajuster automatiquement la méthode de calcul selon l'optimisation
        unite = self.config_vars[mat_name]['unite'].get()
        if unite:  # Si une unité est déjà définie
            nouvelle_methode = self.get_methode_defaut_par_unite_et_opti(unite, optimiser)
            self.config_vars[mat_name]['methode_calcul'].set(nouvelle_methode)
            
            # Log pour informer l'utilisateur
            status = "optimisé" if optimiser else "non optimisé"
            self.log(f"🔧 {mat_name}: {status} → Méthode ajustée: {nouvelle_methode}")
            
            if optimiser:
                self.log(f"   💡 Mode optimisé: Méthodes manuelles pour maîtriser les taux de chute")
            else:
                self.log(f"   💡 Mode non optimisé: API Cadwork pour précision maximale")
            
    def on_config_changed(self, mat_name):
        """Callback quand la configuration d'un matériau change"""
        pass  # Pour l'instant, pas d'action spécifique
        
    def on_price_changed(self, mat_name):
        """Callback quand le prix d'un matériau est modifié"""
        try:
            nouveau_prix = self.safe_float(self.config_vars[mat_name]['prix_unitaire'].get())
            # Validation du prix
            if nouveau_prix < 0:
                self.config_vars[mat_name]['prix_unitaire'].set("0.00")
        except Exception:
            pass
    
    def modify_price(self, mat_name):
        """Permet de modifier le prix d'un matériau"""
        current_price = self.config_vars[mat_name]['prix_unitaire'].get()
        new_price = simpledialog.askfloat(
            "Modification prix", 
            f"Nouveau prix pour {mat_name} (€/unité) :",
            initialvalue=float(current_price),
            minvalue=0.0
        )
        if new_price is not None:
            self.config_vars[mat_name]['prix_unitaire'].set(f"{new_price:.2f}")
            self.log(f"💰 Prix modifié pour {mat_name}: {new_price:.2f}€")
    
    def reset_price(self, mat_name):
        """Remet le prix par défaut depuis Cadwork"""
        prix_cadwork = self.materiaux_detectes[mat_name]['prix_cadwork']
        self.config_vars[mat_name]['prix_unitaire'].set(f"{prix_cadwork:.2f}")
        self.log(f"🔄 Prix resetté pour {mat_name}: {prix_cadwork:.2f}€")
    
    def config_bois_structure(self):
        """Configuration rapide pour bois de structure"""
        for mat_name in self.materiaux_detectes.keys():
            mat_lower = mat_name.lower()
            if any(x in mat_lower for x in ['epicea', 'sapin', 'douglas', 'pin', 'meleze', 'chene', 'kvh', 'gl24', 'bois']):
                self.config_vars[mat_name]['optimiser'].set(True)
                self.config_vars[mat_name]['unite'].set("m³")
                self.on_unite_changed(mat_name)  # Met à jour les méthodes
                self.on_optimiser_changed(mat_name)  # Applique logique optimisé=manuel
        self.log("✅ Configuration 'Bois structure' appliquée (optimisé → méthodes manuelles)")
    
    def config_panneaux(self):
        """Configuration rapide pour panneaux"""
        for mat_name in self.materiaux_detectes.keys():
            mat_lower = mat_name.lower()
            if any(x in mat_lower for x in ['osb', 'panneau', 'contreplaque', 'mdf', 'agglomere']):
                self.config_vars[mat_name]['optimiser'].set(False)
                self.config_vars[mat_name]['unite'].set("m²")
                self.on_unite_changed(mat_name)  # Met à jour les méthodes
                self.on_optimiser_changed(mat_name)  # Applique logique non optimisé=API
        self.log("✅ Configuration 'Panneaux' appliquée (non optimisé → méthodes API)")
    
    def config_enveloppes(self):
        """Configuration rapide pour enveloppes"""
        for mat_name in self.materiaux_detectes.keys():
            mat_lower = mat_name.lower()
            if any(x in mat_lower for x in ['enveloppe', 'paroi', 'isolation', 'bardage']):
                self.config_vars[mat_name]['optimiser'].set(False)
                self.config_vars[mat_name]['unite'].set("m²")
                self.on_unite_changed(mat_name)  # Met à jour les méthodes
                self.on_optimiser_changed(mat_name)  # Applique logique non optimisé=API
        self.log("✅ Configuration 'Enveloppes' appliquée (non optimisé → méthodes API)")
    
    def config_quincaillerie(self):
        """Configuration rapide pour quincaillerie"""
        for mat_name in self.materiaux_detectes.keys():
            mat_lower = mat_name.lower()
            if any(x in mat_lower for x in ['vis', 'boulon', 'clou', 'cheville', 'equerre', 'sabot', 'connecteur', 'acier', 'inox']):
                if any(x in mat_lower for x in ['acier', 'inox', 'metal']):
                    # Métaux au poids
                    self.config_vars[mat_name]['optimiser'].set(False)
                    self.config_vars[mat_name]['unite'].set("Kg")
                else:
                    # Quincaillerie à l'unité
                    self.config_vars[mat_name]['optimiser'].set(False)
                    self.config_vars[mat_name]['unite'].set("U")
                self.on_unite_changed(mat_name)  # Met à jour les méthodes
                self.on_optimiser_changed(mat_name)  # Applique logique non optimisé=API
        self.log("✅ Configuration 'Quincaillerie' appliquée (non optimisé → méthodes API)")
    
    def config_lineaires(self):
        """Configuration rapide pour matériaux linéaires"""
        for mat_name in self.materiaux_detectes.keys():
            mat_lower = mat_name.lower()
            if any(x in mat_lower for x in ['_l', 'barre', 'profile', 'tube', 'rail']):
                self.config_vars[mat_name]['optimiser'].set(True)
                self.config_vars[mat_name]['unite'].set("ml")
                self.on_unite_changed(mat_name)  # Met à jour les méthodes
                self.on_optimiser_changed(mat_name)  # Applique logique optimisé=manuel
        self.log("✅ Configuration 'Linéaires' appliquée (optimisé → méthodes manuelles)")
    
    def config_reset(self):
        """Remet tout à zéro"""
        for mat_name in self.materiaux_detectes.keys():
            self.config_vars[mat_name]['optimiser'].set(False)
            self.config_vars[mat_name]['unite'].set("m³")
            self.on_unite_changed(mat_name)  # Met à jour les méthodes
            self.on_optimiser_changed(mat_name)  # Applique logique non optimisé=API
        self.log("🔄 Configuration remise à zéro (non optimisé → méthodes API)")
    
    def validate_config(self):
        """Valide la configuration et sauvegarde dans les attributs Cadwork"""
        try:
            if not self.materiaux_detectes:
                messagebox.showwarning("Attention", "Aucun matériau à configurer. Effectuez d'abord l'analyse.")
                return
                
            self.log("✅ Validation de la configuration...")
            
            # Sauvegarde dans les attributs Cadwork
            self.save_config_to_cadwork_attributes()
            
            self.config_validee = True
            self.optimize_button.config(state='normal')
            
            # Résumé de la configuration
            self.log("=" * 60)
            self.log("📋 RÉSUMÉ DE LA CONFIGURATION")
            self.log("=" * 60)
            
            count_optimiser = sum(1 for mat_name in self.materiaux_detectes.keys() 
                                if self.config_vars[mat_name]['optimiser'].get())
            count_m3 = sum(1 for mat_name in self.materiaux_detectes.keys() 
                          if self.config_vars[mat_name]['unite'].get() == "m³")
            count_m2 = sum(1 for mat_name in self.materiaux_detectes.keys() 
                          if self.config_vars[mat_name]['unite'].get() == "m²")
            count_ml = sum(1 for mat_name in self.materiaux_detectes.keys() 
                          if self.config_vars[mat_name]['unite'].get() == "ml")
            count_kg = sum(1 for mat_name in self.materiaux_detectes.keys() 
                          if self.config_vars[mat_name]['unite'].get() == "Kg")
            count_u = sum(1 for mat_name in self.materiaux_detectes.keys() 
                         if self.config_vars[mat_name]['unite'].get() == "U")
            
            self.log(f"🔧 {count_optimiser} matériau(x) à optimiser (→ chutes automatiques)")
            self.log(f"📦 {count_m3} matériau(x) en m³ | {count_m2} en m² | {count_ml} en ml")
            self.log(f"⚖️ {count_kg} matériau(x) en Kg | {count_u} en unités")
            self.log("")
            self.log("🎯 LOGIQUE MÉTHODES SELON OPTIMISATION :")
            self.log("   • OPTIMISÉ → Manuel pour maîtriser les calculs et taux de chute")
            self.log("   • NON OPTIMISÉ → API Cadwork pour précision maximale")
            self.log("")
            self.log("📐 CORRESPONDANCES API CADWORK (méthodes cohérentes par unité) :")
            self.log("   🔲 VOLUME (m³) :")
            self.log("     • Volume standard (brut) → gc.get_volume() - Enveloppe géométrique")
            self.log("     • Volume liste → gc.get_list_volume() - Pour nomenclature")
            self.log("     • Volume physique réel → gc.get_actual_physical_volume() - Avec découpes")
            self.log("   📏 SURFACE (m²) :")
            self.log("     • Surface (l×h) → largeur × hauteur manuel")
            self.log("     • Surface face avant → gc.get_area_of_front_face() - API Cadwork")
            self.log("     • Surface face référence → gc.get_element_reference_face_area() - API Cadwork")
            self.log("     • Surface réelle → Volume physique réel / épaisseur")
            self.log("   📐 LONGUEUR (ml) :")
            self.log("     • Manuel longueur → Longueur brute sans calculs")
            self.log("     • Longueur liste → gc.get_list_length() - Pour nomenclature")
            self.log("     • Longueur physique → gc.get_length() - Avec dépassements")
            self.log("   ⚖️ POIDS (Kg) :")
            self.log("     • Manuel poids → Saisie/calcul manuel")
            self.log("     • Poids matériau → mc.get_weight() - Depuis définition matériau")
            self.log("   🔢 QUANTITÉ (U) :")
            self.log("     • Manuel quantité → Saisie manuelle")
            self.log("     • Nombre pièces → Comptage éléments")
            self.log("")
            self.log("💡 Étape suivante : Configuration détaillée de l'optimisation")
            self.log("   (longueurs min/max, pas, marges, priorités...)")
            self.log("=" * 60)
            
            messagebox.showinfo("Configuration validée", 
                              "Configuration de base sauvegardée avec succès !\n\n"
                              "✅ Matériaux à optimiser identifiés\n"
                              "✅ Prix et unités configurés\n"
                              "✅ Types de calculs définis\n\n"
                              "➡️ Étape suivante : Configuration détaillée de l'optimisation\n"
                              "   (longueurs, pas, marges, etc.)")
            
        except Exception as e:
            self.log(f"❌ Erreur validation : {e}")
            messagebox.showerror("Erreur", f"Erreur lors de la validation :\n{e}")
    
    def save_config_to_cadwork_attributes(self):
        """Sauvegarde la configuration dans les attributs Cadwork"""
        try:
            self.log("💾 Sauvegarde configuration dans attributs Cadwork...")
            
            for eid in self.element_ids:
                try:
                    mat_name = ac.get_element_material_name(eid).strip()
                    if mat_name not in self.config_vars:
                        continue
                    
                    config = self.config_vars[mat_name]
                    
                    # Attribut 4: Unité matériau
                    unite = config['unite'].get()
                    ac.set_user_attribute([eid], 4, unite)
                    
                    # Attribut 5: À optimiser (1/0)
                    optimiser = "1" if config['optimiser'].get() else "0"
                    ac.set_user_attribute([eid], 5, optimiser)
                    
                    # Attribut 6: Type optimisation
                    type_opti = config['type_opti'].get()
                    ac.set_user_attribute([eid], 6, type_opti)
                    
                    # Attribut 14: Méthode calcul (conversion nom → code)
                    methode_display = config['methode_calcul'].get()
                    methode_code = self.get_methode_code(methode_display)
                    ac.set_user_attribute([eid], 14, methode_code)
                    
                    # Attribut 17: Taux chute manuel (défaut 0%)
                    ac.set_user_attribute([eid], 17, "0.0")
                    
                    # Attribut 18: Coût traitement chutes (défaut 0€)
                    ac.set_user_attribute([eid], 18, "0.0")
                    
                    # Attribut 19: Prix matière (déjà fait dans save_material_prices_to_attributes)
                    prix_unitaire = self.safe_float(config['prix_unitaire'].get())
                    ac.set_user_attribute([eid], 19, str(prix_unitaire))
                    
                except Exception as e:
                    self.log(f"⚠️ Erreur sauvegarde élément {eid}: {e}")
                    
            self.log("✅ Configuration sauvegardée dans les attributs Cadwork")
            
        except Exception as e:
            raise Exception(f"Erreur sauvegarde configuration : {e}")
    
    def launch_optimization(self):
        """Lance le script d'optimisation"""
        try:
            if not self.config_validee:
                messagebox.showwarning("Configuration requise", 
                                     "Veuillez d'abord valider la configuration.")
                return
                
            self.log("🚀 Lancement de la configuration d'optimisation...")
            
            # Fermeture de l'interface actuelle
            self.root.destroy()
            
            # Lancement du script 2
            script_path = os.path.join(os.path.dirname(__file__), "2_optimisation_scierie.py")
            if os.path.exists(script_path):
                subprocess.Popen([sys.executable, script_path])
            else:
                # Si le script n'existe pas encore, on affiche un message
                messagebox.showinfo("Script suivant", 
                                  "Le script de configuration d'optimisation n'est pas encore créé.\n"
                                  "Configuration de base sauvegardée avec succès.")
                
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du lancement :\n{e}")
    
    def close_application(self):
        """Ferme l'application proprement"""
        if messagebox.askyesno("Confirmation", "Voulez-vous vraiment fermer le configurateur ?"):
            self.root.destroy()

def main():
    """Fonction principale"""
    try:
        # Test de connexion à Cadwork
        _ = ec.get_active_identifiable_element_ids()
        
        # Lancement de l'interface
        app = ConfigurateurDevisInterface()
        app.root.mainloop()
        
    except Exception as e:
        messagebox.showerror("Erreur Cadwork", 
                           f"Impossible de se connecter à Cadwork :\n{e}\n\n"
                           "Assurez-vous que Cadwork est ouvert et qu'un projet est chargé.")

if __name__ == "__main__":
    main()