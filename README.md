# 🌞 Horizon Énergie – Dimensionneur Solaire Sigen

Outil interactif de dimensionnement photovoltaïque pour installations résidentielles utilisant les onduleurs **Sigen Home**.

🔗 **Accès direct à l’outil en ligne :**  
https://dimensionneur-solaire-qvbahekpamth7grjd7wdhw.streamlit.app

---

## 🚀 Fonctionnalités principales

### **Sélection du matériel**
- Panneaux : *Trina*, *Soluxtec* (catalogue intégré)
- Onduleurs : *Sigen Home* monophasés et triphasés 3×400 V
- Batteries : *Sigen 6 kWh* et *10 kWh* (optionnel)

### **Dimensionnement électrique**
- Calcul automatique de la puissance DC totale (Wc)
- Suggestion automatique de l’onduleur selon :
  - type de réseau
  - ratio DC/AC maximal
- Vérification de sécurité :
  - **Voc froid ≤ VDC_max onduleur**
  - **Vmp chaud dans la plage MPPT**

### **Simulation énergétique (Belgique)**
- Production PV mensuelle basée sur 1034 kWh/kWc/an
- 3 profils de consommation mensuelle :
  - Standard  
  - Hiver fort  
  - Été fort  
- 4 profils horaires sur 24 h :
  - Uniforme  
  - Matin + soir  
  - Travail journée  
  - Télétravail  

Calcul automatique :
- Autoconsommation annuelle (kWh)
- Injection réseau (kWh)
- **Taux d’autoconsommation (%)**
- **Taux de couverture (%)**

### **Schéma de câblage**
- Visualisation simple :  
  **Strings → MPPT → Onduleur**
- Compatible 1 ou 2 strings
- Schéma interactif (zoom & pan)

### **Export Excel complet**
Inclut :
- Catalogue matériel
- Paramètres choisis
- Profils production/consommation
- Vérifications électriques (onglet “Strings”)
- Synthèse client

---

## 📁 Structure du projet

```
dimensionneur-solaire/
├── app.py               # Interface Streamlit
├── excel_generator.py   # Génération du fichier Excel
├── logo_horizon.png     # Logo Horizon Énergie
├── requirements.txt     # Dépendances Python
└── README.md            # Documentation
```

---

## 🧠 Notes techniques

Calculs utilisés :

### Voc froid
```
Voc_cold = Ns * Voc * (1 + alpha * (Tmin - 25))
```

### Vmp chaud
```
Vmp_hot = Ns * Vmp * (1 + alpha * (Tmax - 25))
```

Conditions de sécurité :
```
Voc_cold ≤ VDC_max onduleur
VMPP_min ≤ Vmp_hot ≤ VMPP_max
```

Indicateurs énergétiques :
```
Taux autoconsommation = autocons / production PV
Taux de couverture    = production PV / consommation totale
```

---

## ✨ Améliorations possibles
- Simulation batterie (charge/décharge)
- Optimisation automatique du nombre de strings
- Analyse économique (ROI, tarif prosumer)
- Export PDF (lorsque Streamlit Cloud le permet mieux)

