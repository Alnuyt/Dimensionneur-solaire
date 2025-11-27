from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
import datetime


def generate_pdf_bytes(config: dict, summary: dict, logo_path: str = "logo_horizon.png") -> bytes:
    """
    Génère un PDF A4 avec :
    - logo Horizon
    - résumé système (panneaux, onduleur, batterie)
    - production / conso / autoconsommation
    - paramètres de string (N_series, T_min, T_max)
    """

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    # ------------------------------------------------
    # HEADER AVEC LOGO + TITRE
    # ------------------------------------------------
    y_top = height - 15 * mm

    # Logo
    try:
        logo = ImageReader(logo_path)
        c.drawImage(
            logo,
            20 * mm,
            y_top - 15 * mm,
            width=40 * mm,
            preserveAspectRatio=True,
            mask="auto",
        )
    except Exception:
        # Pas de logo, ce n'est pas bloquant
        pass

    c.setFont("Helvetica-Bold", 16)
    c.drawString(70 * mm, y_top, "Étude photovoltaïque")
    c.setFont("Helvetica", 11)
    c.drawString(70 * mm, y_top - 6 * mm, "Horizon Énergie")

    c.setFont("Helvetica", 9)
    today = datetime.date.today().strftime("%d/%m/%Y")
    c.drawString(70 * mm, y_top - 12 * mm, f"Date : {today}")

    # ------------------------------------------------
    # INFOS SYSTÈME
    # ------------------------------------------------
    y = y_top - 25 * mm

    def kv(label, value):
        nonlocal y
        c.setFont("Helvetica-Bold", 10)
        c.drawString(20 * mm, y, f"{label} :")
        c.setFont("Helvetica", 10)
        c.drawString(60 * mm, y, str(value))
        y -= 6 * mm

    c.setFont("Helvetica-Bold", 12)
    c.drawString(20 * mm, y, "Données système")
    y -= 8 * mm

    kv("Panneau", summary.get("panel_id", ""))
    kv("Nombre de panneaux", summary.get("n_modules", ""))
    kv("Puissance unitaire (Wc)", summary.get("p_stc", ""))
    kv("Puissance DC totale (Wc)", summary.get("p_dc_total", ""))
    kv("Type de réseau", summary.get("grid_type", ""))
    kv("Onduleur", summary.get("inverter_id", ""))

    if summary.get("battery_enabled", False):
        kv("Batterie activée", "Oui")
        kv("Capacité batterie (kWh)", summary.get("battery_kwh", ""))
    else:
        kv("Batterie activée", "Non")

    # ------------------------------------------------
    # PERFORMANCES ÉNERGÉTIQUES
    # ------------------------------------------------
    y -= 4 * mm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(20 * mm, y, "Performances énergétiques")
    y -= 8 * mm

    kv("Consommation annuelle (kWh)", round(summary.get("cons_year", 0), 1))
    kv("Production PV annuelle (kWh)", round(summary.get("pv_year", 0), 1))
    kv("Autoconsommation annuelle (kWh)", round(summary.get("autocons_year", 0), 1))
    kv("Taux d'autoconsommation (%)", f"{summary.get('taux_auto', 0):.1f}")
    kv("Taux de couverture (%)", f"{summary.get('taux_couv', 0):.1f}")

    # ------------------------------------------------
    # STRINGS / CONDITIONS ÉLECTRIQUES
    # ------------------------------------------------
    y -= 4 * mm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(20 * mm, y, "Configuration des strings")
    y -= 8 * mm

    kv("Modules en série (string)", summary.get("n_series", ""))
    kv("Température de calcul min (°C)", summary.get("t_min", ""))
    kv("Température de calcul max (°C)", summary.get("t_max", ""))

    y -= 4 * mm
    c.setFont("Helvetica", 9)
    c.drawString(
        20 * mm,
        y,
        "Remarque : les tensions Voc/Vmp sont vérifiées dans le fichier Excel joint (onglet 'Strings').",
    )

    # ------------------------------------------------
    # PIED DE PAGE
    # ------------------------------------------------
    c.setFont("Helvetica", 8)
    c.drawString(
        20 * mm,
        10 * mm,
        "Ce document est une estimation indicative basée sur les données fournies et les profils standard pour la Belgique.",
    )
    c.drawString(
        20 * mm,
        6 * mm,
        "Horizon Énergie – Votre énergie, notre défi.",
    )

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.getvalue()
