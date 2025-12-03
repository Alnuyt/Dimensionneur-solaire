import os
import math
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

from excel_generator import get_catalog, generate_workbook_bytes


# ----------------------------------------------------
# CONFIG STREAMLIT
# ----------------------------------------------------
st.set_page_config(
    page_title="Dimensionneur Solaire Sigen",
    layout="wide",
)

# ----------------------------------------------------
# CATALOGUE
# ----------------------------------------------------
PANELS, INVERTERS, BATTERIES = get_catalog()
PANEL_IDS = [p[0] for p in PANELS]


def get_panel_power(panel_id: str) -> float:
    for p in PANELS:
        if p[0] == panel_id:
            return p[1]
    return 0.0


def get_panel_elec(panel_id: str):
    for p in PANELS:
        if p[0] == panel_id:
            return {
                "id": p[0],
                "Pstc": float(p[1]),
                "Voc": float(p[2]),
                "Vmp": float(p[3]),
                "Isc": float(p[4]),
                "alpha_V": float(p[6]),
            }
    return None


def get_inverter_elec(inv_id: str):
    for inv in INVERTERS:
        if inv[0] == inv_id:
            return {
                "id": inv[0],
                "P_ac": float(inv[1]),
                "P_dc_max": float(inv[2]),
                "Vmpp_min": float(inv[3]),
                "Vmpp_max": float(inv[4]),
                "Vdc_max": float(inv[5]),
                "Impp_max": float(inv[6]),
                "nb_mppt": int(inv[7]),
                "type_reseau": inv[8],
                "famille": inv[9],
            }
    return None


def get_recommended_inverter(p_dc_total, grid_type, max_dc_ac, famille=None):
    for inv in INVERTERS:
        inv_id, p_ac, p_dc_max, _, _, _, _, _, inv_type, inv_family = inv
        if inv_type != grid_type:
            continue
        if famille is not None and inv_family != famille:
            continue
        if p_dc_total <= p_dc_max and p_dc_total / p_ac <= max_dc_ac:
            return inv_id
    return None


# ----------------------------------------------------
# PROFILS
# ----------------------------------------------------

def monthly_pv_profile_kwh_kwp():
    annual_kwh_kwp = 1034.0
    distribution = np.array([3.8, 5.1, 8.7, 11.5, 12.1,
                             11.8, 11.9, 10.8, 9.7, 7.0, 4.3, 3.3])
    return annual_kwh_kwp * distribution / 100.0


def monthly_consumption_profile(annual_kwh, profile):
    profiles = {
        "Standard":   [7, 7, 8, 9, 9, 9, 9, 9, 8, 8, 8, 9],
        "Hiver fort": [10,10,10,9,8,7,6,6,7,8,9,10],
        "Été fort":   [6,6,7,8,9,10,11,11,10,8,7,7],
    }
    arr = np.array(profiles[profile], dtype=float)
    arr = arr / arr.sum()
    return annual_kwh * arr


# ----------------------------------------------------
# OPTIMISATION STRINGS
# ----------------------------------------------------
def optimize_strings(
    N_tot,
    panel,
    inverter,
    T_min,
    T_max,
    ratio_dc_ac_target=1.25,
    ratio_dc_ac_min=1.05,
    ratio_dc_ac_max=1.35,
):
    Voc = panel["Voc"]
    Vmp = panel["Vmp"]
    Isc = panel["Isc"]
    alpha_V = panel["alpha_V"] / 100.0
    Pstc = panel["Pstc"]

    Vdc_max = inverter["Vdc_max"]
    Vmpp_min = inverter["Vmpp_min"]
    Vmpp_max = inverter["Vmpp_max"]
    Impp_max = inverter["Impp_max"]
    nb_mppt = inverter["nb_mppt"]
    P_ac = inverter["P_ac"]

    voc_factor_cold = (1 + alpha_V * (T_min - 25.0))
    vmp_factor_hot = (1 + alpha_V * (T_max - 25.0))

    if voc_factor_cold <= 0 or vmp_factor_hot <= 0:
        return None

    N_series_max = math.floor(Vdc_max / (Voc * voc_factor_cold))
    N_series_min = max(math.ceil(Vmpp_min / (Vmp * vmp_factor_hot)), 6)

    if N_series_min > N_series_max:
        return None

    best = None
    best_score = -1e9

    for N_series in range(N_series_min, N_series_max + 1):

        Voc_cold = N_series * Voc * voc_factor_cold
        Vmp_hot = N_series * Vmp * vmp_factor_hot

        if Voc_cold > Vdc_max:
            continue
        if not (Vmpp_min <= Vmp_hot <= Vmpp_max):
            continue

        N_strings_theo = N_tot // N_series
        if N_strings_theo < 1:
            continue

        N_strings_max_mppt = nb_mppt * 2
        N_strings_max = min(N_strings_theo, N_strings_max_mppt)

        for N_strings in range(1, N_strings_max + 1):
            base = N_strings // nb_mppt
            rest = N_strings % nb_mppt
            strings_per_mppt = [base + (1 if i < rest else 0) for i in range(nb_mppt)]

            for s in strings_per_mppt:
                if s * Isc > Impp_max:
                    break
            else:
                N_used = N_strings * N_series
                P_dc = N_used * Pstc
                ratio = P_dc / P_ac

                if not (ratio_dc_ac_min <= ratio <= ratio_dc_ac_max):
                    continue

                imbalance = max(strings_per_mppt) - min(strings_per_mppt)
                score = (-10 * abs(ratio - ratio_dc_ac_target)
                         + 0.02 * N_used
                         - 5 * (N_strings - 1)
                         - 2 * imbalance
                         + 0.5 * N_series)

                if score > best_score:
                    best_score = score
                    best = {
                        "N_series": N_series,
                        "N_strings": N_strings,
                        "strings_per_mppt": strings_per_mppt,
                        "N_used": N_used,
                        "P_dc": P_dc,
                        "ratio_dc_ac": ratio,
                        "Voc_cold": Voc_cold,
                        "Vmp_hot": Vmp_hot,
                    }

    return best


# ----------------------------------------------------
# SIDEBAR
# ----------------------------------------------------
with st.sidebar:
    st.markdown("### 🔧 Paramètres")

    panel_id = st.selectbox("Panneau", PANEL_IDS)
    n_modules = st.number_input("Nombre de panneaux", 6, 50, 12)

    grid_type = st.selectbox("Type de réseau", ["Mono", "Tri 3x230", "Tri 3x400"])

    store_mode = st.selectbox(
        "Installation compatible SigenStore ?",
        ["Auto", "Oui (Store)", "Non (Hybride)"],
    )

    if store_mode == "Oui (Store)":
        fam_pref = "Store"
    elif store_mode == "Non (Hybride)":
        fam_pref = "Hybride"
    else:
        fam_pref = None

    max_dc_ac = st.slider("Ratio DC/AC max", 1.0, 1.5, 1.30, 0.01)

    battery_enabled = st.checkbox("Batterie", False)
    if battery_enabled:
        battery_kwh = st.slider("Capacité batterie (kWh)", 6.0, 50.0, 6.0, 0.5)
    else:
        battery_kwh = 0.0

    st.markdown("---")
    annual_consumption = st.number_input(
        "Conso annuelle (kWh)", 500, 20000, 3500, 100
    )
    consumption_profile = st.selectbox(
        "Profil mensuel", ["Standard", "Hiver fort", "Été fort"]
    )


    t_min = st.number_input("Temp min (°C)", -30, 10, -10)
    t_max = st.number_input("Temp max (°C)", 30, 90, 70)


# ----------------------------------------------------
# CALCULS PRINCIPAUX
# ----------------------------------------------------
p_stc = get_panel_power(panel_id)
p_dc_total_theo = p_stc * n_modules

recommended = get_recommended_inverter(
    p_dc_total_theo,
    grid_type,
    max_dc_ac,
    fam_pref
)

inv_list = []
if recommended:
    inv_list.append("(Auto) " + recommended)

valid_inverters = [
    inv for inv in INVERTERS
    if inv[8] == grid_type and (fam_pref is None or inv[9] == fam_pref)
]

if not valid_inverters:
    valid_inverters = [inv for inv in INVERTERS if inv[8] == grid_type]

inv_list += [inv[0] for inv in valid_inverters]

chosen = st.sidebar.selectbox("Onduleur", inv_list)

if chosen.startswith("(Auto) "):
    inverter_id = recommended
else:
    inverter_id = chosen

panel_elec = get_panel_elec(panel_id)
inv_elec = get_inverter_elec(inverter_id)

opt_result = optimize_strings(
    n_modules, panel_elec, inv_elec, t_min, t_max
)

if opt_result:
    N_used = opt_result["N_used"]
    P_dc = opt_result["P_dc"]
    ratio = opt_result["ratio_dc_ac"]
else:
    N_used = n_modules
    P_dc = p_dc_total_theo
    ratio = P_dc / inv_elec["P_ac"]

p_dc_kwp = P_dc / 1000.0

pv_profile = monthly_pv_profile_kwh_kwp() * p_dc_kwp
cons_profile = monthly_consumption_profile(annual_consumption, consumption_profile)
autocons = np.minimum(pv_profile, cons_profile)

pv_year = pv_profile.sum()
autocons_year = autocons.sum()

# ----------------------------------------------------
# AFFICHAGE
# ----------------------------------------------------
st.title("Dimensionneur Solaire Sigen – Horizon Énergie")

col1, col2, col3 = st.columns(3)

with col1:
    st.metric("Puissance DC théorique", f"{p_dc_total_theo:.0f} Wc")
    st.metric("Puissance DC câblée", f"{P_dc:.0f} Wc")

with col2:
    st.metric("Production annuelle", f"{pv_year:.0f} kWh")
    st.metric("Autocons.", f"{100*autocons_year/pv_year:.1f} %")

with col3:
    st.metric("Onduleur", inverter_id)
    st.metric("Ratio DC/AC", f"{ratio:.2f}")

# ----------------------------------------------------
# GRAPHIQUES
# ----------------------------------------------------
st.markdown("## 📊 Profil mensuel")

df = pd.DataFrame({
    "Mois": ["Jan", "Fév", "Mar", "Avr", "Mai", "Juin",
             "Juil", "Août", "Sep", "Oct", "Nov", "Déc"],
    "Consommation": cons_profile,
    "Production PV": pv_profile,
})

fig = px.bar(df, x="Mois", y=["Consommation", "Production PV"])
st.plotly_chart(fig, use_container_width=True)

# ----------------------------------------------------
# PROFIL HORAIRE – JOUR TYPE
# ----------------------------------------------------
st.markdown("## 🕒 Profil horaire – jour type")

days_in_month = np.array([31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31])
idx = month_for_hours - 1

day_cons = cons_monthly[idx] / days_in_month[idx]
day_pv = pv_monthly[idx] / days_in_month[idx]

cons_frac = get_hourly_profile(horaire_profile)
cons_hour = day_cons * cons_frac

pv_frac = np.array([
    0, 0, 0, 0, 0,
    0.01, 0.04, 0.07, 0.10, 0.13, 0.14, 0.14,
    0.13, 0.10, 0.07, 0.04, 0.02,
    0, 0, 0, 0, 0, 0, 0,
])
if pv_frac.sum() > 0:
    pv_frac = pv_frac / pv_frac.sum()
pv_hour = day_pv * pv_frac

autocons_hour = np.minimum(cons_hour, pv_hour)

hours = np.arange(24)
df_hour = pd.DataFrame({
    "Heure": hours,
    "Consommation (kWh)": cons_hour,
    "Production PV (kWh)": pv_hour,
    "Autoconsommation (kWh)": autocons_hour,
})

fig2 = px.line(
    df_hour,
    x="Heure",
    y=["Consommation (kWh)", "Production PV (kWh)", "Autoconsommation (kWh)"],
    markers=True,
    labels={"value": "kWh", "variable": ""},
    color_discrete_sequence=["#E74C3C", "#F1C40F", "#2ECC71"]  # rouge / jaune / vert
)
st.plotly_chart(fig2, use_container_width=True)
st.dataframe(df_hour)

# ----------------------------------------------------
# EXPORT EXCEL
# ----------------------------------------------------
st.markdown("## 📥 Export Excel")

cfg = {
    "panel_id": panel_id,
    "n_modules": int(n_modules),
    "grid_type": grid_type,
    "battery_enabled": battery_enabled,
    "battery_kwh": battery_kwh,
    "max_dc_ac": max_dc_ac,
    "annual_consumption": annual_consumption,
    "consumption_profile": consumption_profile,
    "t_min": t_min,
    "t_max": t_max,
    "n_series": opt_result["N_series"] if opt_result else n_modules,
    "inverter_id": inverter_id,
}

if st.button("Générer l’Excel"):
    file = generate_workbook_bytes(cfg)
    st.download_button(
        "Télécharger",
        data=file,
        file_name="Dimensionnement_Sigen.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
