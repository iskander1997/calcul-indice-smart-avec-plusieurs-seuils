import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os

class ExcelFileSelector:
    def __init__(self, master=None, callback=None):
        # Création de la fenêtre principale
        self.callback = callback
        self.root = tk.Tk() if master is None else tk.Toplevel(master)
        self.root.title("Sélection des Fichiers Excel")
        self.root.geometry("800x600")
        
        # Palette de couleurs moderne
        self.colors = {
            'background': 'ghost white',
            'primary': 'blue4',
            'secondary': 'blue4',
            'accent': '#2c3e50',
            'text': 'gray16',
            'white': '#ffffff'
        }
        
        # Configuration du style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Configuration de la fenêtre
        self.root.configure(bg=self.colors['background'])
        self.root.resizable(False, False)
        
        # Variables de stockage
        self.prices_path = tk.StringVar()
        self.scores_path = tk.StringVar()
        self.thresholds = tk.StringVar()
        
        # Création du cadre principal
        self.create_ui()

    def create_ui(self):
        # Cadre principal
        main_frame = tk.Frame(self.root, bg=self.colors['background'], padx=40, pady=30)
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Titre
        title_label = tk.Label(main_frame, text="Sélection des Fichiers Excel", 
                                font=('Segoe UI', 20, 'bold'), 
                                fg=self.colors['accent'], 
                                bg=self.colors['background'])
        title_label.pack(pady=(0, 30))
        
        # Fichier des Prix
        # self.create_file_selector(main_frame, 
        #                           "Historique des Cours", 
        #                           self.prices_path, 
        #                           self.select_prices_file)
        
        # Fichier des Scores
        self.create_file_selector(main_frame, 
                                  "Scores des Sociétés", 
                                  self.scores_path, 
                                  self.select_scores_file)
        
        # Section Seuils
        self.create_thresholds_section(main_frame)
        
        # Bouton de Validation
        validate_button = tk.Button(main_frame, 
                                    text="Valider", 
                                    command=self.validate_inputs,
                                    bg=self.colors['primary'], 
                                    fg=self.colors['white'],
                                    font=('Segoe UI', 12, 'bold'),
                                    relief=tk.FLAT,
                                    activebackground=self.colors['secondary'])
        validate_button.pack(pady=(30, 0), ipadx=20, ipady=10)

    def create_file_selector(self, parent, label_text, path_var, select_command):
        # Cadre pour chaque sélecteur de fichier
        frame = tk.Frame(parent, bg=self.colors['background'])
        frame.pack(fill=tk.X, pady=10)
        
        # Label
        label = tk.Label(frame, 
                         text=label_text, 
                         font=('Segoe UI', 12), 
                         fg=self.colors['text'], 
                         bg=self.colors['background'])
        label.pack(anchor='w')
        
        # Sous-cadre pour l'entrée et les boutons
        input_frame = tk.Frame(frame, bg=self.colors['background'])
        input_frame.pack(fill=tk.X, pady=(5, 0))
        
        # Entrée de fichier
        entry = tk.Entry(input_frame, 
                         textvariable=path_var, 
                         font=('Segoe UI', 10), 
                         width=60, 
                         state='readonly',
                         relief=tk.FLAT,
                         bg=self.colors['white'])
        entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 10))
        
        # Bouton Parcourir
        browse_btn = tk.Button(input_frame, 
                               text="Parcourir", 
                               command=select_command,
                               bg=self.colors['secondary'], 
                               fg=self.colors['white'],
                               font=('Segoe UI', 10),
                               relief=tk.FLAT,
                               activebackground=self.colors['primary'])
        browse_btn.pack(side=tk.LEFT)
        
        # Bouton Effacer
        clear_btn = tk.Button(input_frame, 
                              text="×", 
                              command=lambda: path_var.set(''),
                              bg=self.colors['white'], 
                              fg=self.colors['accent'],
                              font=('Segoe UI', 10, 'bold'),
                              width=3,
                              relief=tk.FLAT)
        clear_btn.pack(side=tk.LEFT, padx=(10, 0))

    def create_thresholds_section(self, parent):
        # Cadre pour les seuils
        frame = tk.Frame(parent, bg=self.colors['background'])
        frame.pack(fill=tk.X, pady=10)
        
        # Label
        label = tk.Label(frame, 
                         text="Seuils (1-199, séparés par des virgules)", 
                         font=('Segoe UI', 12), 
                         fg=self.colors['text'], 
                         bg=self.colors['background'])
        label.pack(anchor='w')
        
        # Entrée des seuils
        entry = tk.Entry(frame, 
                         textvariable=self.thresholds, 
                         font=('Segoe UI', 10), 
                         relief=tk.FLAT,
                         bg=self.colors['white'])
        entry.pack(fill=tk.X, pady=(5, 0))

    # def select_prices_file(self):
    #     filepath = filedialog.askopenfilename(
    #         title="Sélectionner le fichier Excel des prix",
    #         filetypes=[("Fichiers Excel", "*.xlsx *.xls")]
    #     )
    #     if filepath:
    #         self.prices_path.set(filepath)

    def select_scores_file(self):
        filepath = filedialog.askopenfilename(
            title="Sélectionner le fichier Excel des scores",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls")]
        )
        if filepath:
            self.scores_path.set(filepath)
            self.prices_path.set(filepath)


    def validate_inputs(self):
        # Validation des fichiers
        # if not self.prices_path.get():
        #     messagebox.showerror("Erreur", "Veuillez sélectionner le fichier des prix")
        #     return None
        
        if not self.scores_path.get():
            messagebox.showerror("Erreur", "Veuillez sélectionner le fichier des scores")
            return None
        
        # Validation des seuils
        thresholds_str = self.thresholds.get().strip()
        if not thresholds_str:
            messagebox.showerror("Erreur", "Veuillez entrer au moins un seuil")
            return None
        
        try:
            thresholds = [float(x.strip()) for x in thresholds_str.split(',')]
            
            # Vérification que les seuils sont entre 1 et 199
            if not all(1 <= x <= 199 for x in thresholds):
                raise ValueError("Tous les seuils doivent être entre 1 et 199")
            
            # Préparation du résultat
            result = {
                'prices_path': self.prices_path.get(),
                'scores_path': self.scores_path.get(),
                'thresholds': thresholds
            }
            
            messagebox.showinfo("Succès", "Tous les fichiers et seuils sont validés!")
            
            # Fermer la fenêtre si un callback est défini
            if self.callback:
                self.callback(result)
            
            self.root.quit()
            return result
            
        except ValueError as e:
            messagebox.showerror("Erreur", str(e))
            return None

    def get_paths(self):
        return self.validation_result

    def run(self):
        #retour de valeurs
        self.root.mainloop()
        return self.validate_inputs()


# if __name__ == "__main__":
#     app = ExcelFileSelector()
#     app.run()

if __name__ == "__main__":
    def on_validate(data):
        print("Données validées :", data)

    app = ExcelFileSelector(callback=on_validate)
    app.run()