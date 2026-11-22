#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Sep 10 13:27:19 2026

@author: stamura
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jul  8 09:53:37 2026

@author: stamura
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jul  6 13:55:20 2026

@author: stamura
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jun 17 11:13:26 2026

@author: stamura
"""
import tkinter as tk
from tkinter import messagebox
from pathlib import Path
import pandas as pd
import numpy as np
import urllib
import matplotlib.pyplot as plt

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.patches import Circle
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import Figure

from scipy import stats
from scipy.ndimage import gaussian_filter

from mpl_toolkits.mplot3d import Axes3D

# for google sheets



# for excel
import openpyxl
from openpyxl import Workbook

import os

import plotly.graph_objects as go
import plotly.io as pio
import plotly.express as px
import webbrowser

pio.renderers.default = "browser"




try:
    from pygame import mixer
except ImportError:
    mixer = None

import mplcursors

import os

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
ASSETS_DIR = BASE_DIR / "assets"


###########################################################
# SYNTHETIC TELESCOPE / CCD SETTINGS
###########################################################

# Size of the simulated detector image
DETECTOR_SIZE = 700

# Actual telescope fields of view
FINDER_FOV = 6.0
MAIN_SCOPE_FOV = 1.2

# How much sky is generated around each LSS structure
SYNTHETIC_FIELD_RADIUS = 18.0

# Number of synthetic foreground/background stars
N_BACKGROUND_STARS = 30000

# Detector appearance
SKY_BACKGROUND = 18.0
READ_NOISE = 1.8

# How far the telescope moves while a direction button is held
FINDER_MOVE_STEP = 0.10
MAIN_SCOPE_MOVE_STEP = 0.015

# Pixel distance from detector center required to select galaxy
FINDER_TARGET_RADIUS = 16
MAIN_TARGET_RADIUS = 22

# Useful while WE are developing.
# True = show known galaxy locations
# False = student only sees the simulated telescope image
DEBUG_CATALOG_OVERLAY = False

# Stores fixed synthetic skies once they are generated
SYNTHETIC_STAR_FIELDS = {}

# Fixed detector noise pattern.
# Keeping this constant stops the screen from flickering while moving.
_detector_rng = np.random.default_rng(2026)

DETECTOR_NOISE_PATTERN = _detector_rng.normal(
    0,
    1,
    (DETECTOR_SIZE, DETECTOR_SIZE)
)

# Persistent Matplotlib telescope canvas
skyfig = None
sky_ax = None
skyplot = None

# for galaxy finder btns and stuff


APP_BG = "#050b12"          # almost black
PANEL_BG = "#e5e7eb"        # soft light gray
CARD_BG = "#f8fafc"         # card background
CARD_BORDER = "#9ca3af"
TEXT_DARK = "#0f172a"
TEXT_BLUE = "#0b3b75"
TEXT_GREEN = "#15803d"

BUTTON_BG = "#f8fafc"
PRIMARY_BUTTON_BG = "#d9f7d8"
PRIMARY_BUTTON_ACTIVE = "#baf0b8"

TITLE_FONT = ("Times New Roman", 17, "bold")
HEADER_FONT = ("Times New Roman", 20, "bold")
BUTTON_FONT = ("Times New Roman", 12, "bold")
SMALL_FONT = ("Times New Roman", 10)
SMALL_BOLD = ("Times New Roman", 10, "bold")



###########################################################
# sound file
class SilentSound:
    def play(self, *args, **kwargs): pass
    def stop(self): pass

def load_sound(filename):
    try:
        if mixer is None: return SilentSound()
        if not mixer.get_init(): mixer.init()
        return mixer.Sound(str(ASSETS_DIR / filename))
    except Exception:
        return SilentSound()

move_sound = load_sound("telescope.wav")
dome_sound = load_sound("dome.wav")

# import real galaxies
# Catalog positions are real; the detector background and spectra are simulated.
CATALOG_FILE = DATA_DIR / "clean_master_galaxy_catalog.csv"
if not CATALOG_FILE.exists():
    CATALOG_FILE = DATA_DIR / "sample_galaxy_catalog.csv"
if not CATALOG_FILE.exists():
    raise FileNotFoundError("Place sample_galaxy_catalog.csv in the data/ folder.")
REAL_CATALOG = pd.read_csv(CATALOG_FILE, dtype={"Galaxy_ID": str})
required = ["Galaxy_ID", "RA_deg", "Dec_deg", "Redshift"]
if any(col not in REAL_CATALOG for col in required):
    raise ValueError("Catalog requires Galaxy_ID, RA_deg, Dec_deg, Redshift")
for col in required[1:]:
    REAL_CATALOG[col] = pd.to_numeric(REAL_CATALOG[col], errors="coerce")
REAL_CATALOG = REAL_CATALOG[np.isfinite(REAL_CATALOG[["RA_deg", "Dec_deg"]]).all(axis=1)
    & REAL_CATALOG["Dec_deg"].between(-90, 90)].copy()
REAL_CATALOG["RA_deg"] %= 360
REAL_CATALOG = REAL_CATALOG.drop_duplicates(subset=["Galaxy_ID", "RA_deg", "Dec_deg"]).reset_index(drop=True)
# The core exercise uses the nearby Hubble approximation, not a cosmological distance model.
MAX_LAB_REDSHIFT = 0.10
OBSERVABLE_CATALOG = REAL_CATALOG[REAL_CATALOG["Redshift"].between(0, MAX_LAB_REDSHIFT, inclusive="right")].copy()
structure_file = DATA_DIR / "known_large_scale_structures.csv"
STRUCTURES = pd.read_csv(structure_file) if structure_file.exists() else pd.DataFrame(columns=["Name", "RA_deg", "Dec_deg"])
ACTIVE_STRUCTURE = None
ACTIVE_STRUCTURE_CATALOG = OBSERVABLE_CATALOG

from scipy.spatial import cKDTree

def unit_vectors(ra, dec):
    ra, dec = np.radians(ra), np.radians(dec)
    return np.column_stack([np.cos(dec)*np.cos(ra), np.cos(dec)*np.sin(ra), np.sin(dec)])

GALAXY_TREE = cKDTree(unit_vectors(OBSERVABLE_CATALOG.RA_deg, OBSERVABLE_CATALOG.Dec_deg))

def nearby_galaxies(ra, dec, radius):
    center = unit_vectors([ra], [dec])[0]
    indices = GALAXY_TREE.query_ball_point(center, 2*np.sin(np.radians(radius)/2))
    return OBSERVABLE_CATALOG.iloc[sorted(indices)].copy()

C_LIGHT = 299792.458
H0 = 70

CA_H_LAB = 396.885
CA_K_LAB = 393.367


# create file

EXCEL_FILE = str(BASE_DIR / "galaxy_database.xlsx")

def setup_excel_file():
    if not os.path.exists(EXCEL_FILE):
        workbook = Workbook()

        spectral_sheet = workbook.active
        spectral_sheet.title = "Spectral Data"
        spectral_sheet.append([
            "Galaxy",
            "Redshift",
            "Velocity (km/s)",
            "Distance (Mpc)",
            "RA",
            "DEC",
            "Calcium H",
            "Calcium K"])

        phot_sheet = workbook.create_sheet(title="Photometry Data")
        phot_sheet.append([
            "Galaxy",
            "Magnitude",
            "Photon Counts",
            "RA",
            "DEC"])

        wedge_sheet = workbook.create_sheet(title="Wedge Map Data")
        wedge_sheet.append([
            "Galaxy",
            "Distance",
            "RA",
            "DEC",
            "X",
            "Y",
            "Z"])

        workbook.save(EXCEL_FILE)

if not os.path.exists(EXCEL_FILE):
    setup_excel_file()

###########################################################
# Telescope coordinates
ra_center = 0
dec_center = 0

# Scope mode
scope_mode = "Finder"

# Observation mode
observation_mode = "Spectral"

# Slit position
slit_x = 0
slit_y = 0

spec_window = None 

galaxy_data = []

gmm_plot_open = False

 
###########################################################
# TELESCOPE MOVEMENT
###########################################################

def get_current_fov():

    if scope_mode == "Finder":
        return FINDER_FOV

    return MAIN_SCOPE_FOV


def get_move_step():

    if scope_mode == "Finder":
        return FINDER_MOVE_STEP

    return MAIN_SCOPE_MOVE_STEP


def redraw_spectrum_if_open():

    try:
        draw_spectrum(
            ax,
            canvas,
            exposure_slider.get()
        )

    except Exception:
        pass


def update_coordinate_labels():

    try:
        ra_label.config(
            text=f"RA Center: {ra_center:.3f}°"
        )

        dec_label.config(
            text=f"Dec Center: {dec_center:.3f}°"
        )

    except Exception:
        pass


def move_north():

    global dec_center

    dec_center += get_move_step()

    # Declination physically cannot exceed ±90 degrees
    dec_center = min(dec_center, 90.0)

    update_coordinate_labels()
    draw_sky()
    redraw_spectrum_if_open()


def move_south():

    global dec_center

    dec_center -= get_move_step()

    dec_center = max(dec_center, -90.0)

    update_coordinate_labels()
    draw_sky()
    redraw_spectrum_if_open()


def move_east():

    global ra_center

    ra_center += get_move_step()

    # Keep RA between 0 and 360 degrees
    ra_center %= 360.0

    update_coordinate_labels()
    draw_sky()
    redraw_spectrum_if_open()


def move_west():

    global ra_center

    ra_center -= get_move_step()

    ra_center %= 360.0

    update_coordinate_labels()
    draw_sky()
    redraw_spectrum_if_open()

##################################
###########################################################
# SYNTHETIC COORDINATE-ACCURATE TELESCOPE SKY
###########################################################


def wrapped_ra_difference(ra, center_ra):
    """
    Return the shortest RA difference from center_ra.

    This prevents problems near RA = 0 / 360 degrees.
    """

    ra = np.asarray(ra, dtype=float)

    return (
        (ra - center_ra + 180.0) % 360.0
    ) - 180.0


###########################################################

def generate_star_field():
    """Deterministic synthetic stars across the sphere; query only the current field."""
    if "all_sky" not in SYNTHETIC_STAR_FIELDS:
        rng = np.random.default_rng(2026)
        n = 600000
        ra = rng.uniform(0, 360, n)
        dec = np.degrees(np.arcsin(rng.uniform(-1, 1, n)))
        SYNTHETIC_STAR_FIELDS["all_sky"] = dict(ra=ra, dec=dec,
            flux=np.clip(rng.lognormal(4, 1.15, n), 8, 2500),
            tree=cKDTree(unit_vectors(ra, dec)))
    field = SYNTHETIC_STAR_FIELDS["all_sky"]
    radius = get_current_fov()
    indices = field["tree"].query_ball_point(unit_vectors([ra_center], [dec_center])[0],
                                           2*np.sin(np.radians(radius)/2))
    return {key: field[key][indices] for key in ("ra", "dec", "flux")}


def sky_to_pixel(ra, dec, center_ra, center_dec, fov):
    """Gnomonic sky projection, valid across RA=0 and at celestial poles."""
    ra, dec = np.radians(np.asarray(ra, dtype=float)), np.radians(np.asarray(dec, dtype=float))
    ra0, dec0 = np.radians([center_ra, center_dec])
    dra = ra-ra0
    denominator = np.sin(dec0)*np.sin(dec) + np.cos(dec0)*np.cos(dec)*np.cos(dra)
    with np.errstate(divide="ignore", invalid="ignore"):
        east = np.cos(dec)*np.sin(dra)/denominator
        north = (np.cos(dec0)*np.sin(dec)-np.sin(dec0)*np.cos(dec)*np.cos(dra))/denominator
    scale = DETECTOR_SIZE / (2*np.tan(np.radians(fov)/2))
    x, y = DETECTOR_SIZE/2-east*scale, DETECTOR_SIZE/2+north*scale
    return np.where(denominator>0, x, np.nan), np.where(denominator>0, y, np.nan)


def add_sources_to_layer(layer, x, y, flux):
    """
    Deposit source brightness into detector pixels.
    """

    x = np.asarray(x)
    y = np.asarray(y)
    flux = np.asarray(flux)

    inside = (
        (x >= 0) &
        (x < DETECTOR_SIZE) &
        (y >= 0) &
        (y < DETECTOR_SIZE)
    )

    if not np.any(inside):
        return

    xi = np.rint(x[inside]).astype(int)
    yi = np.rint(y[inside]).astype(int)

    xi = np.clip(
        xi,
        0,
        DETECTOR_SIZE - 1
    )

    yi = np.clip(
        yi,
        0,
        DETECTOR_SIZE - 1
    )

    np.add.at(
        layer,
        (yi, xi),
        flux[inside]
    )


###########################################################

def create_ccd_image():
    """
    Construct the simulated telescope detector image.

    Returns:
        image
        visible galaxy dataframe
        galaxy x pixels
        galaxy y pixels
    """

    fov = get_current_fov()

    #######################################################
    # Empty detector layers
    #######################################################

    star_layer = np.zeros(
        (DETECTOR_SIZE, DETECTOR_SIZE),
        dtype=float
    )

    galaxy_layer = np.zeros(
        (DETECTOR_SIZE, DETECTOR_SIZE),
        dtype=float
    )

    #######################################################
    # SYNTHETIC STARS
    #######################################################

    star_field = generate_star_field()

    if star_field is not None:

        star_ra = star_field["ra"]
        star_dec = star_field["dec"]
        star_flux = star_field["flux"]

        star_x, star_y = sky_to_pixel(
            star_ra,
            star_dec,
            ra_center,
            dec_center,
            fov
        )

        add_sources_to_layer(
            star_layer,
            star_x,
            star_y,
            star_flux
        )

    #######################################################
    # REAL GALAXIES
    #######################################################

    galaxy_df = nearby_galaxies(ra_center, dec_center, fov)

    galaxy_x, galaxy_y = sky_to_pixel(
        galaxy_df["RA_deg"].to_numpy(),
        galaxy_df["Dec_deg"].to_numpy(),
        ra_center,
        dec_center,
        fov
    )

    galaxy_visible = (
        (galaxy_x >= 0) &
        (galaxy_x < DETECTOR_SIZE) &
        (galaxy_y >= 0) &
        (galaxy_y < DETECTOR_SIZE)
    )

    visible_galaxies = galaxy_df.loc[
        galaxy_visible
    ].copy()

    visible_x = galaxy_x[
        galaxy_visible
    ]

    visible_y = galaxy_y[
        galaxy_visible
    ]

    #######################################################
    # Give galaxies realistic approximate image brightness.
    #
    # Redshift affects apparent brightness slightly,
    # but the actual science redshift remains untouched.
    #######################################################

    if len(visible_galaxies) > 0:

        if "Redshift" in visible_galaxies.columns:

            galaxy_z = pd.to_numeric(
                visible_galaxies["Redshift"],
                errors="coerce"
            ).fillna(0.05).to_numpy()

        else:

            galaxy_z = np.full(
                len(visible_galaxies),
                0.05
            )

        galaxy_flux = (
            500
            +
            1600 /
            (
                1
                +
                18 *
                np.clip(galaxy_z, 0, None)
            )
        )

        add_sources_to_layer(
            galaxy_layer,
            visible_x,
            visible_y,
            galaxy_flux
        )

    #######################################################
    # Point Spread Functions
    #
    # Stars = compact
    # Galaxies = more extended
    #######################################################

    if scope_mode == "Finder":

        star_sigma = 0.9
        galaxy_sigma = 2.2

    else:

        star_sigma = 1.2
        galaxy_sigma = 4.0

    star_layer = gaussian_filter(
        star_layer,
        sigma=star_sigma
    )

    galaxy_layer = gaussian_filter(
        galaxy_layer,
        sigma=galaxy_sigma
    )

    #######################################################
    # Combine detector components
    #######################################################

    image = (
        SKY_BACKGROUND
        +
        star_layer
        +
        galaxy_layer
        +
        READ_NOISE * DETECTOR_NOISE_PATTERN
    )

    image = np.clip(
        image,
        0,
        None
    )

    #######################################################
    # FITS-like logarithmic stretch
    #######################################################

    display_image = np.log1p(
        image
    )

    return (
        display_image,
        visible_galaxies,
        visible_x,
        visible_y
    )


###########################################################
# DRAW TELESCOPE SKY
###########################################################

def draw_sky():

    global current_target
    global skyfig
    global sky_ax
    global skyplot

    if ACTIVE_STRUCTURE is None:
        return

    if ACTIVE_STRUCTURE_CATALOG is None:
        return

    #######################################################
    # Create the Matplotlib canvas ONLY ONCE.
    #
    # Your old version destroyed and rebuilt the canvas
    # every movement step, which contributed to jumpiness.
    #######################################################

    if skyplot is None:

        skyfig = Figure(
            figsize=(10, 8),
            dpi=100
        )

        skyfig.patch.set_facecolor(
            "black"
        )

        sky_ax = skyfig.add_subplot(111)

        skyplot = FigureCanvasTkAgg(
            skyfig,
            master=skyframe
        )

        skyplot.get_tk_widget().pack(
            fill="both",
            expand=True
        )

    #######################################################
    # Generate current CCD exposure
    #######################################################

    (
        image,
        visible_galaxies,
        galaxy_x,
        galaxy_y
    ) = create_ccd_image()

    sky_ax.clear()

    sky_ax.set_facecolor(
        "black"
    )

    #######################################################
    # Display detector
    #######################################################

    vmin = np.percentile(
        image,
        5
    )

    vmax = np.percentile(
        image,
        99.85
    )

    if vmax <= vmin:
        vmax = image.max()

    sky_ax.imshow(
        image,
        origin="lower",
        cmap="gray",
        vmin=vmin,
        vmax=vmax,
        interpolation="nearest"
    )

    #######################################################
    # Find galaxy nearest center of telescope
    #######################################################

    current_target = None

    detector_center = (
        DETECTOR_SIZE / 2
    )

    if len(visible_galaxies) > 0:

        pixel_distances = np.sqrt(
            (galaxy_x - detector_center) ** 2
            +
            (galaxy_y - detector_center) ** 2
        )

        nearest_position = np.argmin(
            pixel_distances
        )

        nearest_pixel_distance = pixel_distances[
            nearest_position
        ]

        if scope_mode == "Finder":

            target_radius = FINDER_TARGET_RADIUS

        else:

            target_radius = MAIN_TARGET_RADIUS

        if nearest_pixel_distance <= target_radius:

            current_target = visible_galaxies.iloc[
                nearest_position
            ]

            target_x = galaxy_x[
                nearest_position
            ]

            target_y = galaxy_y[
                nearest_position
            ]

            ###################################################
            # Yellow confirmation ring
            ###################################################

            target_circle = Circle(
                (target_x, target_y),
                radius=target_radius + 6,
                fill=False,
                edgecolor="yellow",
                linewidth=2.0,
                alpha=0.9,
                zorder=20
            )

            sky_ax.add_patch(
                target_circle
            )

    #######################################################
    # DEBUG GALAXY LOCATIONS
    #######################################################

    if (
        DEBUG_CATALOG_OVERLAY
        and len(visible_galaxies) > 0
    ):

        sky_ax.scatter(
            galaxy_x,
            galaxy_y,
            s=40,
            facecolors="none",
            edgecolors="cyan",
            linewidths=0.7,
            alpha=0.7,
            zorder=15
        )

    #######################################################
    # TELESCOPE RETICLE
    #######################################################

    cx = DETECTOR_SIZE / 2
    cy = DETECTOR_SIZE / 2

    if scope_mode == "Finder":

        ###################################################
        # Finder crosshair
        ###################################################

        sky_ax.plot(
            [cx - 35, cx - 8],
            [cy, cy],
            color="#38bdf8",
            linewidth=1.4,
            zorder=30
        )

        sky_ax.plot(
            [cx + 8, cx + 35],
            [cy, cy],
            color="#38bdf8",
            linewidth=1.4,
            zorder=30
        )

        sky_ax.plot(
            [cx, cx],
            [cy - 35, cy - 8],
            color="#38bdf8",
            linewidth=1.4,
            zorder=30
        )

        sky_ax.plot(
            [cx, cx],
            [cy + 8, cy + 35],
            color="#38bdf8",
            linewidth=1.4,
            zorder=30
        )

        finder_circle = Circle(
            (cx, cy),
            radius=8,
            fill=False,
            edgecolor="#38bdf8",
            linewidth=1.3,
            zorder=30
        )

        sky_ax.add_patch(
            finder_circle
        )

    else:

        ###################################################
        # Main telescope spectrograph slit
        ###################################################

        slit_half_width = 5
        slit_half_height = 80

        sky_ax.plot(
            [cx - slit_half_width,
             cx - slit_half_width],
            [cy - slit_half_height,
             cy + slit_half_height],
            color="#38bdf8",
            linewidth=1.5,
            zorder=30
        )

        sky_ax.plot(
            [cx + slit_half_width,
             cx + slit_half_width],
            [cy - slit_half_height,
             cy + slit_half_height],
            color="#38bdf8",
            linewidth=1.5,
            zorder=30
        )

        sky_ax.plot(
            [cx - 18, cx + 18],
            [cy, cy],
            color="#38bdf8",
            linewidth=1.0,
            zorder=30
        )

    #######################################################
    # DOME SHUTTER
    #######################################################

    opening = dome_offset / 100.0

    cover = 0.5 * (
        1 - opening
    )

    if cover > 0:

        sky_ax.add_patch(
            plt.Rectangle(
                (0, 1 - cover),
                1,
                cover,
                transform=sky_ax.transAxes,
                color="#5f6670",
                alpha=0.97,
                zorder=100
            )
        )

        sky_ax.add_patch(
            plt.Rectangle(
                (0, 0),
                1,
                cover,
                transform=sky_ax.transAxes,
                color="#5f6670",
                alpha=0.97,
                zorder=100
            )
        )

    #######################################################
    # Clean telescope appearance
    #######################################################

    sky_ax.set_xlim(
        0,
        DETECTOR_SIZE
    )

    sky_ax.set_ylim(
        0,
        DETECTOR_SIZE
    )

    sky_ax.set_aspect(
        "equal"
    )

    sky_ax.axis(
        "off"
    )

    skyfig.subplots_adjust(
        left=0,
        right=1,
        top=1,
        bottom=0
    )

    #######################################################
    # Status
    #######################################################

    try:

        if current_target is None:

            readinglbl.config(
                text="No centered lab target. Blank fields may reflect missing survey coverage."
            )

        else:

            readinglbl.config(
                text=
                f"Galaxy {current_target['Galaxy_ID']} centered."
            )

    except Exception:
        pass

    #######################################################
    # Redraw existing canvas
    #######################################################

    skyplot.draw_idle()

###########################################################
# window 2: the spectral acquisition
def open_spectrum_window():

    global spec_window

    if spec_window is not None and spec_window.winfo_exists():
        spec_window.lift()
        redraw_spectrum_if_open()
        return

    spec_window = tk.Toplevel(window)

    spec_window.title(f"{observation_mode} Acquisition")
    
    spec_window.geometry("1200x900+50+50")

    spec_window.minsize(900, 650)

    spec_window.resizable(True, True)
    
    spec_window.lift()
    
    build_spectrum_ui(spec_window)
    
############################################################
# Window 3: Redshift Table Window: Wedge graph 

galaxy_data = []

def open_gmm_wedge_plot(auto_update=False):
    """Compatibility callback: the core map shows student measurements, without GMM."""
    df = load_spectral_observations()
    if df.empty:
        if not auto_update:
            messagebox.showinfo("3D Wedge Plot", "Measure and record at least one galaxy first.")
        return
    fields = ["Redshift", "RA", "DEC"]
    for field in fields:
        df[field] = pd.to_numeric(df[field], errors="coerce")
    df = df.dropna(subset=fields)
    df = df[np.isfinite(df[fields]).all(axis=1) & (df["Redshift"] > 0)]
    if df.empty:
        messagebox.showinfo("3D Wedge Plot", "No valid positive-redshift measurements are available.")
        return
    distance = C_LIGHT * df["Redshift"].to_numpy() / H0
    ra, dec = np.radians(df["RA"]), np.radians(df["DEC"])
    x = distance * np.cos(dec) * np.cos(ra)
    y = distance * np.cos(dec) * np.sin(ra)
    z = distance * np.sin(dec)
    details = np.column_stack([df["Galaxy"].astype(str), df["RA"], df["DEC"], df["Redshift"], distance])
    fig = go.Figure(go.Scatter3d(
        x=x, y=y, z=z, mode="markers", name="Student measurements",
        marker=dict(size=5, color="#38bdf8", opacity=0.85), customdata=details,
        hovertemplate="Galaxy %{customdata[0]}<br>RA %{customdata[1]}°<br>Dec %{customdata[2]}°<br>Measured z %{customdata[3]}<br>Distance %{customdata[4]} Mpc<extra></extra>"))
    fig.add_trace(go.Scatter3d(x=[0], y=[0], z=[0], mode="markers", name="Observer",
                             marker=dict(size=7, color="#fbbf24", symbol="diamond")))
    fig.update_layout(title="3D Wedge Plot — Student-Measured Redshifts",
        template="plotly_dark", scene=dict(xaxis_title="X (Mpc)", yaxis_title="Y (Mpc)",
        zaxis_title="Z (Mpc)", aspectmode="data"))
    html_file = BASE_DIR / "student_3d_wedge_plot.html"
    fig.write_html(str(html_file), include_plotlyjs=True)
    if not auto_update:
        webbrowser.open(html_file.resolve().as_uri())


def add_spectral_reading():
    """All spectral recording must pass through the student measurement analyzer."""
    save_and_measure_spectrum()


def add_photometry_reading():

    galaxy_name = f"P-{np.random.randint(100,999)}"

    apparent_magnitude = round(
        np.random.uniform(10, 18), 2)

    brightness = round(
        np.random.uniform(1000, 5000), 1)

    ra = round(ra_center, 2)

    dec = round(dec_center, 2)

    workbook = openpyxl.load_workbook(EXCEL_FILE)

    sheet = workbook["Photometry Data"]

    sheet.append([
        galaxy_name,
        apparent_magnitude,
        brightness,
        ra,
        dec])

    workbook.save(EXCEL_FILE)

    if "acquisition_status_label" in globals():
        try:
            acquisition_status_label.config(
            text=f"{galaxy_name} photometry saved.",
            fg="#15803d")
        except tk.TclError:
            pass
    
####################
# WEDGE PLOT
def update_wedge_plot():

    wedge_ax.clear()

    wedge_ax.set_facecolor("black")

    xvals = []
    yvals = []
    zvals = []

    for galaxy in galaxy_data:

        r = galaxy["distance"]

        theta = np.radians(galaxy["ra"])

        x = r * np.cos(np.radians(galaxy["dec"])) * np.cos(theta)

        y = r * np.cos(np.radians(galaxy["dec"])) * np.sin(theta)

        z = r * np.sin(np.radians(galaxy["dec"]))

        xvals.append(x)
        yvals.append(y)
        zvals.append(z)

    if len(xvals) > 0:

        colors = ['cyan'] * len(xvals)

        colors[-1] = 'red'

        wedge_ax.scatter(
            xvals,
            yvals,
            zvals,
            s=70,
            c=colors,
            depthshade=True)

    wedge_ax.view_init(
        elev=25,
        azim=45)

    wedge_ax.set_title(
        "3D Galaxy Redshift Map",
        color='white',
        fontsize=14)

    wedge_ax.set_xlabel("X (Mpc)")
    wedge_ax.set_ylabel("Y (Mpc)")
    wedge_ax.set_zlabel("Z (Mpc)")

    wedge_ax.tick_params(colors='white')

    wedge_ax.xaxis.label.set_color('white')
    wedge_ax.yaxis.label.set_color('white')
    wedge_ax.zaxis.label.set_color('white')

    wedge_ax.grid(color='gray', alpha=0.3)

    wedge_ax.autoscale_view()

    wedge_canvas.draw()
###########################################################
# control map with buttons

dome_open = False
dome_offset = 0


def toggle_dome():

    global dome_open

    dome_open = not dome_open
    
    dome_status_label.config(
    text=f"Dome: {'Open' if dome_open else 'Closed'}",
    fg="#22c55e" if dome_open else "#f87171")

    try:
        dome_sound.play()
    except:
        pass

    domebtn.config(text="Dome: Open" if dome_open else "Dome: Closed")

    if dome_open:
        animate_dome_open()
    else:
        animate_dome_close()


def animate_dome_open():

    global dome_offset

    if dome_offset < 100:
        dome_offset += 5
        draw_sky()
        window.after(25, animate_dome_open)


def animate_dome_close():

    global dome_offset

    if dome_offset > 0:
        dome_offset -= 5
        draw_sky()
        window.after(25, animate_dome_close)
    else:
        dome_offset = 0
        draw_sky()


def switch_view():

    global scope_mode

    if scope_mode == "Finder":

        scope_mode = "Main Scope"

    else:

        scope_mode = "Finder"

    viewbtn.config(
        text=f"View: {scope_mode}")

    draw_sky()


def switch_mode():
    global observation_mode
    observation_mode = "Photometry" if observation_mode == "Spectral" else "Spectral"
    for name in ("modebtn", "spectrum_modebtn"):
        widget = globals().get(name)
        if widget is not None and widget.winfo_exists():
            widget.config(text=f"Mode: {observation_mode}")
    widget = globals().get("acquisitionbtn")
    if widget is not None and widget.winfo_exists():
        widget.config(text=f"Open {observation_mode} Acquisition")
    redraw_spectrum_if_open()


def set_coordinates():
    global ra_center, dec_center
    try:
        ra, dec = float(ra_entry.get()), float(dec_entry.get())
        if not np.isfinite([ra, dec]).all() or not -90 <= dec <= 90:
            raise ValueError
    except ValueError:
        messagebox.showerror("Coordinates", "Enter numeric RA and Dec in degrees; Dec must be between -90 and 90.")
        return
    ra_center, dec_center = ra % 360, dec
    update_coordinate_labels()
    draw_sky()
    redraw_spectrum_if_open()


def take_reading():
    """From the Galaxy Finder, send the selected target to the acquisition window."""

    if observation_mode == "Spectral":
        if current_target is None:
            try:
                readinglbl.config(
                    text="No galaxy selected — center a galaxy in the reticle first.")
            except Exception:
                pass
            return

        # A telescope reading begins the observing workflow.  The spectrum is
        # acquired and saved from the Spectral Acquisition window itself.
        open_spectrum_window()

    else:
        add_photometry_reading()

###########################################################
def make_multi_spectrum(redshift, exposure, brightness=1200.0):
    """Create a simulated galaxy spectrum. Redshift is simulation truth only."""
    wavelength = np.linspace(380, 700, 3201)
    continuum = float(brightness) * np.ones_like(wavelength)

    lines = [
        CA_K_LAB * (1 + redshift),  # Calcium K
        CA_H_LAB * (1 + redshift),  # Calcium H
        486.1 * (1 + redshift)]     # H-beta

    true_spectrum = continuum.copy()
    for line in lines:
        true_spectrum -= 450 * np.exp(
            -((wavelength - line) ** 2) / (2 * 0.65 ** 2))

    noise_level = 400 / np.sqrt(max(float(exposure), 1.0))
    observed = true_spectrum + np.random.normal(
        0, noise_level, len(wavelength))

    return wavelength, true_spectrum, observed


# The spectrum currently displayed in the acquisition window and the last
# observation explicitly saved by the student.  Neither contains catalog z.
current_spectrum_observation = None
saved_spectrum = None


def save_current_spectrum():
    global saved_spectrum

    if current_target is None:
        if "acquisition_status_label" in globals():
            acquisition_status_label.config(
                text="No galaxy selected. Center a galaxy first.",
                fg="#b91c1c")
        return False

    if (current_spectrum_observation is None or
            str(current_spectrum_observation["Galaxy_ID"]) != str(current_target["Galaxy_ID"])):
        if "acquisition_status_label" in globals():
            acquisition_status_label.config(
                text="No spectrum is available to save yet.",
                fg="#b91c1c")
        return False

    # Copy the exact noisy spectrum the student is looking at.
    saved_spectrum = {
        "Galaxy_ID": current_spectrum_observation["Galaxy_ID"],
        "RA_deg": current_spectrum_observation["RA_deg"],
        "Dec_deg": current_spectrum_observation["Dec_deg"],
        "Exposure": current_spectrum_observation["Exposure"],
        "Wavelength": current_spectrum_observation["Wavelength"].copy(),
        "Observed": current_spectrum_observation["Observed"].copy()
    }

    if "acquisition_status_label" in globals():
        acquisition_status_label.config(
            text=f"Spectrum saved for Galaxy {saved_spectrum['Galaxy_ID']}.",
            fg="#15803d")

    print("Spectrum observation saved:")
    print("Galaxy:", saved_spectrum["Galaxy_ID"])
    print("Exposure:", saved_spectrum["Exposure"])
    return True


def update_spectrum_object_card():

    if 'spectrum_obj_id_label' not in globals():
        return

    if current_target is not None:
        obj_id = str(current_target["Galaxy_ID"])
        obj_ra = f"{float(current_target['RA_deg']):.2f}°"
        obj_dec = f"{float(current_target['Dec_deg']):.2f}°"
        obj_z = "Not measured"
        obj_v = "Not measured"
    else:
        obj_id = "----"
        obj_ra = "----"
        obj_dec = "----"
        obj_z = "Not measured"
        obj_v = "Not measured"

    spectrum_obj_id_label.config(text=f"Galaxy ID: {obj_id}")
    spectrum_obj_ra_label.config(text=f"RA: {obj_ra}")
    spectrum_obj_dec_label.config(text=f"Dec: {obj_dec}")
    spectrum_obj_z_label.config(text=f"Redshift: {obj_z}")
    spectrum_obj_v_label.config(text=f"Velocity: {obj_v}")

##############################################
# DRAW SPECTRUM

def draw_spectrum(ax, canvas, exposure):
    global current_spectrum_observation

    ax.clear()
    update_spectrum_object_card()
    ax.set_facecolor("black")

    if observation_mode == "Spectral":
        brightness = 1200.0

        if current_target is not None and not pd.isna(current_target["Redshift"]):
            # Catalog redshift is used ONLY to simulate where the lines appear.
            hidden_redshift = float(current_target["Redshift"])
            wavelength, true_spectrum, observed = make_multi_spectrum(
                hidden_redshift, exposure, brightness=brightness)

            current_spectrum_observation = {
                "Galaxy_ID": str(current_target["Galaxy_ID"]),
                "RA_deg": float(current_target["RA_deg"]),
                "Dec_deg": float(current_target["Dec_deg"]),
                "Exposure": float(exposure),
                "Wavelength": wavelength.copy(),
                "Observed": observed.copy()
            }
        else:
            wavelength = np.linspace(380, 700, 3201)
            true_spectrum = brightness * np.ones_like(wavelength)
            observed = true_spectrum + np.random.normal(
                0, 400 / np.sqrt(max(float(exposure), 1.0)), len(wavelength))
            current_spectrum_observation = None

        scatter = ax.scatter(
            wavelength, observed, s=8, color='white', alpha=0.7,
            label="Photon Measurements")

        cursor = mplcursors.cursor(scatter, hover=True)

        @cursor.connect("add")
        def on_add(sel):
            x, y = sel.target
            sel.annotation.set_text(
                f"Wavelength: {x:.2f} nm\nIntensity: {y:.0f}")
            sel.annotation.arrow_patch.set(color='white')

        ax.set_title("Spectral Acquisition", color='white')
        ax.set_xlabel("Wavelength (nm)", color='white', fontsize=12)
        ax.set_ylabel("Intensity", color='white', fontsize=12)
        ax.legend()

        photon_count = int(exposure * brightness)
        signal_noise = round(
            np.mean(observed) / max(np.std(observed - true_spectrum), 1e-9), 1)

        info_text = (
            f"Object: Galaxy\n"
            f"Exposure Time: {exposure:.1f} s\n"
            f"Photon Count: {photon_count}\n"
            f"Signal/Noise: {signal_noise}")

        ax.text(
            0.02, 0.97, info_text, transform=ax.transAxes, fontsize=9,
            color='white', verticalalignment='top',
            bbox=dict(facecolor='black', alpha=0.7, edgecolor='white'))

    else:
        current_spectrum_observation = None
        times = np.linspace(0, 10, 200)
        brightness = (
            1000 + 100 * np.sin(times) +
            np.random.normal(0, 30, len(times)))

        ax.plot(times, brightness, color='yellow', linewidth=2)
        ax.scatter(times, brightness, s=8, color='#38bdf8')

        distance_factor = np.sqrt(ra_center**2 + dec_center**2)
        magnitude = round(10 + distance_factor / 50, 2)
        photon_counts = int(np.mean(brightness))
        variability = round(np.std(brightness), 1)

        info_text = (
            f"Object: Galaxy\n"
            f"Magnitude: {magnitude}\n"
            f"Photon Counts: {photon_counts}\n"
            f"Variability: {variability}\n"
            f"Observation Mode: Photometry")

        ax.text(
            0.02, 0.97, info_text, transform=ax.transAxes, fontsize=9,
            color='white', verticalalignment='top',
            bbox=dict(facecolor='black', alpha=0.7, edgecolor='white'))
        ax.set_title("Photometry", color='white')
        ax.set_xlabel("Time", color='white', fontsize=12)
        ax.set_ylabel("Brightness", color='white', fontsize=12)

    ax.tick_params(colors='white')
    canvas.draw()

def save_and_measure_spectrum():
    """Save the spectrum currently displayed, then open the Spectrum Analyzer."""
    if save_current_spectrum():
        open_record_measurements_window()


##############################
# SPECTRUM ANALYZER / DATA TABLE

def open_record_measurements_window():
    """Measure Ca K and Ca H from the student's saved noisy spectrum."""

    if saved_spectrum is None:
        if "acquisition_status_label" in globals():
            acquisition_status_label.config(
                text="Save a spectrum before opening the analyzer.",
                fg="#b91c1c")
        return

    observation = {key: (value.copy() if isinstance(value, np.ndarray) else value)
                   for key, value in saved_spectrum.items()}
    recorded = False
    record_window = tk.Toplevel(window)
    record_window.title("Spectrum Analyzer")
    record_window.geometry("1000x720+100+80")
    record_window.configure(bg=APP_BG)

    tk.Label(
        record_window,
        text="SPECTRUM ANALYZER",
        font=HEADER_FONT, bg=APP_BG, fg="white"
    ).pack(pady=(12, 4))

    tk.Label(
        record_window,
        text=(f"Galaxy {observation['Galaxy_ID']}   |   "
              f"RA {observation['RA_deg']:.2f}°   |   "
              f"Dec {observation['Dec_deg']:.2f}°   |   "
              f"Exposure {observation['Exposure']:.0f} s"),
        font=SMALL_FONT, bg=APP_BG, fg="#e2e8f0"
    ).pack(pady=(0, 8))

    tk.Label(
        record_window,
        text="Click the observed Ca K minimum first, then the Ca H minimum. The program will calculate z from YOUR measurements.",
        font=SMALL_BOLD, bg=APP_BG, fg="#f8fafc"
    ).pack(pady=(0, 8))

    fig = Figure(figsize=(9, 4.5), dpi=100, facecolor=APP_BG)
    measure_ax = fig.add_subplot(111)
    measure_ax.set_facecolor("black")
    measure_ax.scatter(
        observation["Wavelength"], observation["Observed"],
        s=10, color="white", alpha=0.75)
    measure_ax.set_title("Saved Observed Spectrum", color="white")
    measure_ax.set_xlabel("Wavelength (nm)", color="white")
    measure_ax.set_ylabel("Intensity", color="white")
    measure_ax.tick_params(colors="white")

    measure_canvas = FigureCanvasTkAgg(fig, master=record_window)
    measure_ax.set_xlim(385, 470)
    from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
    toolbar = NavigationToolbar2Tk(measure_canvas, record_window, pack_toolbar=False)
    toolbar.update()
    toolbar.pack(fill="x")
    measure_canvas.draw()
    measure_canvas.get_tk_widget().pack(fill="both", expand=True, padx=12, pady=6)

    controls = tk.Frame(record_window, bg=APP_BG)
    controls.pack(fill="x", padx=12, pady=8)

    tk.Label(controls, text="Ca K measured (nm):", bg=APP_BG, fg="white").grid(row=0, column=0, sticky="e", padx=5)
    k_entry = tk.Entry(controls, width=12)
    k_entry.grid(row=0, column=1, padx=5)

    tk.Label(controls, text="Ca H measured (nm):", bg=APP_BG, fg="white").grid(row=0, column=2, sticky="e", padx=5)
    h_entry = tk.Entry(controls, width=12)
    h_entry.grid(row=0, column=3, padx=5)

    result_label = tk.Label(
        controls, text="Waiting for Ca K measurement...",
        justify="left", bg=APP_BG, fg="#f8fafc", font=SMALL_FONT)
    result_label.grid(row=1, column=0, columnspan=5, sticky="w", pady=8)

    selected_lines = []

    def clear_measurements():
        for artist in selected_lines:
            try:
                artist.remove()
            except Exception:
                pass
        selected_lines.clear()
        k_entry.delete(0, tk.END)
        h_entry.delete(0, tk.END)
        result_label.config(text="Waiting for Ca K measurement...")
        measure_canvas.draw_idle()

    def on_spectrum_click(event):
        if toolbar.mode or event.button != 1 or event.inaxes != measure_ax or event.xdata is None:
            return

        # Snap to the nearest actual wavelength channel.
        wavelengths = observation["Wavelength"]
        idx = int(np.argmin(np.abs(wavelengths - event.xdata)))
        measured = float(wavelengths[idx])

        if not k_entry.get():
            k_entry.insert(0, f"{measured:.3f}")
            line = measure_ax.axvline(measured, color="cyan", linestyle="--", linewidth=1.4, label="Measured Ca K")
            selected_lines.append(line)
            result_label.config(text="Ca K recorded. Now click the Ca H minimum.")
        elif not h_entry.get():
            h_entry.insert(0, f"{measured:.3f}")
            line = measure_ax.axvline(measured, color="yellow", linestyle="--", linewidth=1.4, label="Measured Ca H")
            selected_lines.append(line)
            result_label.config(text="Both lines selected. Click Calculate / Record.")
        else:
            result_label.config(text="Both measurements are already set. Use Reset Measurements to try again.")

        measure_canvas.draw_idle()

    measure_canvas.mpl_connect("button_press_event", on_spectrum_click)

    def calculate_and_save():
        nonlocal recorded
        if recorded:
            result_label.config(text="This observation is already recorded. Acquire a new spectrum to repeat it.")
            return
        try:
            k_measured = float(k_entry.get())
            h_measured = float(h_entry.get())
        except ValueError:
            result_label.config(text="Measure both Ca K and Ca H before recording.")
            return

        if not np.isfinite([k_measured, h_measured]).all() or not (380 <= k_measured < h_measured <= 700):
            result_label.config(text="Enter finite wavelengths within 380–700 nm, with Ca K shorter than Ca H.")
            return
        z_k = (k_measured - CA_K_LAB) / CA_K_LAB
        z_h = (h_measured - CA_H_LAB) / CA_H_LAB
        redshift = (z_k + z_h) / 2.0
        if redshift <= 0 or abs(z_k - z_h) > 0.002:
            result_label.config(text="Check the calcium pair: positive average z is required and z(K)/z(H) must agree within 0.002.")
            return
        velocity = C_LIGHT * redshift
        distance = velocity / H0

        result_label.config(
            text=(f"z(K) = {z_k:.5f}    z(H) = {z_h:.5f}    "
                  f"Average z = {redshift:.5f}    "
                  f"v = {velocity:.0f} km/s    d = {distance:.1f} Mpc"))

        if not save_manual_spectral_reading(
            observation["Galaxy_ID"], redshift, velocity, distance,
            h_measured, k_measured,
            ra=observation["RA_deg"], dec=observation["Dec_deg"]):
            result_label.config(text="Could not save. Close the workbook in Excel and try again.")
            return
        recorded = True

        if "acquisition_status_label" in globals():
            acquisition_status_label.config(
                text=f"Measured Galaxy {observation['Galaxy_ID']} | z = {redshift:.5f}",
                fg="#15803d")

    make_button(controls, "Reset Measurements", clear_measurements).grid(row=0, column=4, padx=8)
    make_button(controls, "Calculate / Record", calculate_and_save, primary=True).grid(row=2, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
    make_button(controls, "Close Analyzer", record_window.destroy).grid(row=2, column=3, columnspan=2, sticky="ew", padx=5, pady=5)

##########################################
def save_manual_spectral_reading(
        galaxy_name,
        redshift,
        velocity,
        distance,
        calcium_h,
        calcium_k,
        ra=None,
        dec=None):

    global galaxy_data

    if ra is None:
        ra = ra_center
    if dec is None:
        dec = dec_center
    ra = float(ra)
    dec = float(dec)

    new_row = (
        f"{galaxy_name:<10}"
        f"{redshift:<10.5f}"
        f"{int(velocity):<15}"
        f"{distance:<10.1f}")

    try:
        workbook = openpyxl.load_workbook(EXCEL_FILE)
    except (OSError, ValueError):
        return False

    spectral_sheet = workbook["Spectral Data"]
    spectral_sheet.append([
        galaxy_name,
        redshift,
        float(velocity),
        float(distance),
        ra,
        dec,
        calcium_h,
        calcium_k])

    ra_rad = np.radians(ra)
    dec_rad = np.radians(dec)

    x = distance * np.cos(dec_rad) * np.cos(ra_rad)
    y = distance * np.cos(dec_rad) * np.sin(ra_rad)
    z = distance * np.sin(dec_rad)

    wedge_sheet = workbook["Wedge Map Data"]
    wedge_sheet.append([
        galaxy_name,
        float(distance),
        ra,
        dec,
        x,
        y,
        z])

    try:
        temporary = str(Path(EXCEL_FILE).with_suffix(".tmp.xlsx"))
        workbook.save(temporary)
        os.replace(temporary, EXCEL_FILE)
    except OSError:
        workbook.close()
        return False
    workbook.close()
    galaxy_data.append({
        "text": new_row,
        "redshift": redshift,
        "velocity": velocity,
        "distance": distance,
        "ra": ra,
        "dec": dec,
        "galaxy": galaxy_name,
"mode": "Single",})
    
    if "db_tree" in globals():
        refresh_observation_database()


    if 'wedge_ax' in globals() and 'wedge_canvas' in globals():
        update_wedge_plot()

    # Update the acquisition window status, if it is open
    if "acquisition_status_label" in globals():
        try:
            acquisition_status_label.config(
            text=f"Saved Galaxy {galaxy_name}  |  z = {redshift:.5f}",
            fg="#15803d")
        except tk.TclError:
            pass
#######################################
# MOS Window

    return True

def open_multi_object_window():

    multi_window = tk.Toplevel(window)
    multi_window.title("Multi-Object Spectroscopy")
    multi_window.geometry("1300x850+70+70")
    multi_window.configure(bg=APP_BG)

    selected_vars = []

    field_size = 3

    field_df = nearby_galaxies(ra_center, dec_center, field_size)

    

    if len(field_df) > 40:
        field_df = field_df.sample(40, random_state=42)

    # MAIN LAYOUT
    left_frame = tk.Frame(multi_window, bg="black")
    left_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

    right_frame = tk.Frame(multi_window, bg=PANEL_BG, width=380)
    right_frame.pack(side="right", fill="y", padx=10, pady=10)
    right_frame.pack_propagate(False)

    # TITLE
    tk.Label(
        right_frame,
        text="MULTI-OBJECT\nSPECTROSCOPY",
        font=HEADER_FONT,
        bg=PANEL_BG,
        fg=TEXT_DARK).pack(pady=(10, 15))

    info_box = make_card(right_frame, "FIELD INFO")
    card_label(info_box, "RA Center:", f"{ra_center:.2f}°", TEXT_BLUE)
    card_label(info_box, "Dec Center:", f"{dec_center:.2f}°", TEXT_BLUE)
    card_label(info_box, "Targets in Field:", str(len(field_df)), TEXT_BLUE)

    # FIELD / FIBER MASK
    fig = Figure(figsize=(8, 6), dpi=100, facecolor=APP_BG)
    field_ax = fig.add_subplot(111)
    field_ax.set_facecolor("black")

    fiber_rows = list(field_df.iterrows())

    fiber_rows = list(field_df.iterrows())

    def redraw_fiber_field():
        field_ax.clear()
        field_ax.set_facecolor("black")

        for fiber_num, (_, row) in enumerate(fiber_rows, start=1):

            selected = False

            if fiber_num <= len(selected_vars):
                selected = selected_vars[fiber_num - 1][0].get() == 1

            edge_color = "#22c55e" if selected else "cyan"
            text_color = "#22c55e" if selected else "white"
            size = 90 if selected else 55
            line_width = 2.2 if selected else 1.2

            field_ax.scatter(
                row["RA_deg"],
                row["Dec_deg"],
                s=size,
                facecolors="none",
                edgecolors=edge_color,
                linewidths=line_width,
                alpha=0.95,
                zorder=3)

            field_ax.text(
                row["RA_deg"] + 0.04,
                row["Dec_deg"] + 0.04,
                str(fiber_num),
                color=text_color,
                fontsize=8,
                zorder=4)

            if selected:
                field_ax.plot(
                    [ra_center, row["RA_deg"]],
                    [dec_center, row["Dec_deg"]],
                    color="#22c55e",
                    linewidth=0.7,
                    alpha=0.55,
                    zorder=2)

        field_ax.scatter(
            ra_center,
            dec_center,
            marker="+",
            s=180,
            color="yellow",
            linewidths=2,
            zorder=5)

        field_ax.set_xlim(ra_center - field_size, ra_center + field_size)
        field_ax.set_ylim(dec_center - field_size, dec_center + field_size)
        field_ax.set_xticks([])
        field_ax.set_yticks([])
        field_ax.set_title("Fiber Mask Field of View", color="white")

        for spine in field_ax.spines.values():
            spine.set_visible(False)

        field_canvas.draw()

    field_canvas = FigureCanvasTkAgg(fig, master=left_frame)
    field_canvas.get_tk_widget().pack(fill="both", expand=True)
    # TARGET TABLE
    table_box = make_card(right_frame, "FIBER TARGETS")

    canvas_table = tk.Canvas(table_box, bg=PANEL_BG, height=260, highlightthickness=0)
    scrollbar = tk.Scrollbar(table_box, orient="vertical", command=canvas_table.yview)
    scroll_frame = tk.Frame(canvas_table, bg=PANEL_BG)

    scroll_frame.bind(
        "<Configure>",
        lambda event: canvas_table.configure(scrollregion=canvas_table.bbox("all")))
    

    canvas_table.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas_table.configure(yscrollcommand=scrollbar.set)

    canvas_table.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    fiber_counter = 1

    for _, row in field_df.iterrows():
        var = tk.IntVar()
        selected_vars.append((var, row))

        cb = tk.Checkbutton(
            scroll_frame,
            text=f"Fiber {fiber_counter:02d} | {row['Galaxy_ID']} | RA {row['RA_deg']:.2f} | Dec {row['Dec_deg']:.2f}",
            variable=var,
            command=redraw_fiber_field,
            bg=PANEL_BG,
            fg=TEXT_DARK,
            font=SMALL_FONT,
            anchor="w",
            justify="left")
        cb.pack(fill="x", pady=2)

        fiber_counter += 1
        
    redraw_fiber_field()

    # EXPOSURE
    exposure_box = make_card(right_frame, "EXPOSURE")

    exposure_slider_multi = tk.Scale(
        exposure_box,
        from_=1,
        to=100,
        orient="horizontal",
        label="Exposure Time (seconds)",
        bg=PANEL_BG,
        font=SMALL_FONT)
    exposure_slider_multi.set(20)
    exposure_slider_multi.pack(fill="x")

    result_label = tk.Label(
        exposure_box,
        text="No spectra acquired yet.",
        bg=PANEL_BG,
        fg=TEXT_DARK,
        font=SMALL_FONT,
        justify="left")
    result_label.pack(fill="x", pady=8)
    
    # EMBEDDED STACKED SPECTRUM
    spectrum_fig = Figure(figsize=(8, 3), dpi=100)
    spectrum_ax = spectrum_fig.add_subplot(111)
    spectrum_ax.set_facecolor("black")
    spectrum_ax.set_title("Stacked Spectrum", color="white")
    spectrum_ax.tick_params(colors="white")

    spectrum_canvas = FigureCanvasTkAgg(
    spectrum_fig,
    master=left_frame)
    
    spectrum_canvas.draw()
    spectrum_canvas.get_tk_widget().pack(fill="x", pady=8)

    def acquire_multi_spectra():

        selected = [
            row for var, row in selected_vars
            if var.get() == 1]

        if len(selected) == 0:
            result_label.config(text="Select at least one galaxy.")
            return

        exposure = exposure_slider_multi.get()

        all_observed = []
        redshifts = []

        for row in selected:
            redshift = float(row["Redshift"])
            redshifts.append(redshift)

            wavelength, true_spectrum, observed = make_multi_spectrum(
                redshift,
                exposure)

            all_observed.append(observed)

        mean_spectrum = np.mean(all_observed, axis=0)

        signal_noise = round(
            np.mean(mean_spectrum) / np.std(mean_spectrum - np.mean(all_observed, axis=0)),
            2)

        spectrum_ax.clear()
        spectrum_ax.set_facecolor("black")

        spectrum_ax.plot(
            wavelength,
            mean_spectrum,
            color="yellow",
            linewidth=1.8,
            label=f"{len(selected)} fibers stacked")
        

        spectrum_ax.set_title("Stacked Multi-Object Spectrum", color="white")
        spectrum_ax.set_xlabel("Wavelength (nm)", color="white")
        spectrum_ax.set_ylabel("Intensity", color="white")
        spectrum_ax.tick_params(colors="white")
        spectrum_ax.legend()

        spectrum_canvas.draw()

        result_label.config(
            text=
            f"Acquired spectra for {len(selected)} galaxies.\n"
            f"Exposure: {exposure} s\n"
            f"Mean z: {np.mean(redshifts):.5f}")
        

    acquire_btn = make_button(
        right_frame,
        "Acquire Selected Spectra",
        acquire_multi_spectra,
        primary=True)
    acquire_btn.pack(fill="x", padx=10, pady=12)

    close_btn = make_button(
        right_frame,
        "Close MOS Window",
        multi_window.destroy)
    close_btn.pack(fill="x", padx=10, pady=(0, 12))
#######################################################
# spectral acquisition window 2

def build_spectrum_ui(spec_window):

    global ax
    global canvas
    global exposure_slider
    global spectrum_modebtn

    spec_window.title(f"LSS Observatory — {observation_mode} Acquisition")
    spec_window.geometry("1300x850+70+70")
    spec_window.minsize(1100, 720)
    spec_window.configure(bg=APP_BG)

    # MAIN LAYOUT
    graphframe = tk.Frame(
        spec_window,
        bg="black",
        highlightbackground="#334155",
        highlightthickness=2)
    graphframe.pack(side="left", fill="both", expand=True, padx=10, pady=10)

    controlframe = tk.Frame(
        spec_window,
        bg=PANEL_BG,
        width=380)
    controlframe.pack(side="right", fill="y", padx=10, pady=10)
    controlframe.pack_propagate(False)

    # TITLE
    tk.Label(
        controlframe,
        text=f"{observation_mode.upper()}\nACQUISITION",
        font=HEADER_FONT,
        bg=PANEL_BG,
        fg=TEXT_DARK).pack(pady=(15, 12))

    # CURRENT OBJECT CARD
    object_box = make_card(controlframe, "CURRENT OBJECT")

    if current_target is not None:
        obj_id = str(current_target["Galaxy_ID"])
        obj_ra = f"{float(current_target['RA_deg']):.2f}°"
        obj_dec = f"{float(current_target['Dec_deg']):.2f}°"
        obj_z = "Not measured"
        obj_v = "Not measured"
    else:
        obj_id = "----"
        obj_ra = "----"
        obj_dec = "----"
        obj_z = "----"
        obj_v = "----"

    global spectrum_obj_id_label
    global spectrum_obj_ra_label
    global spectrum_obj_dec_label
    global spectrum_obj_z_label
    global spectrum_obj_v_label

    spectrum_obj_id_label = tk.Label(object_box, text=f"Galaxy ID: {obj_id}", bg=PANEL_BG, fg=TEXT_BLUE, font=SMALL_BOLD, anchor="w")
    spectrum_obj_id_label.pack(fill="x", pady=2)
    
    spectrum_obj_ra_label = tk.Label(object_box, text=f"RA: {obj_ra}", bg=PANEL_BG, fg=TEXT_BLUE, font=SMALL_BOLD, anchor="w")
    spectrum_obj_ra_label.pack(fill="x", pady=2)

    spectrum_obj_dec_label = tk.Label(object_box, text=f"Dec: {obj_dec}", bg=PANEL_BG, fg=TEXT_BLUE, font=SMALL_BOLD, anchor="w")
    spectrum_obj_dec_label.pack(fill="x", pady=2)

    spectrum_obj_z_label = tk.Label(object_box, text=f"Redshift: {obj_z}", bg=PANEL_BG, fg=TEXT_BLUE, font=SMALL_BOLD, anchor="w")
    spectrum_obj_z_label.pack(fill="x", pady=2)

    spectrum_obj_v_label = tk.Label(object_box, text=f"Velocity: {obj_v}", bg=PANEL_BG, fg=TEXT_BLUE, font=SMALL_BOLD, anchor="w")
    spectrum_obj_v_label.pack(fill="x", pady=2)

    # FIGURE
    fig = Figure(figsize=(8, 6), dpi=100, facecolor=APP_BG)
    ax = fig.add_subplot(111)
    ax.set_facecolor("black")

    canvas = FigureCanvasTkAgg(fig, master=graphframe)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True)

    # EXPOSURE CARD
    exposure_box = make_card(controlframe, "EXPOSURE")

    exposure_slider = tk.Scale(
        exposure_box,
        from_=1,
        to=100,
        orient="horizontal",
        label="Exposure Time (seconds)",
        bg=PANEL_BG,
        font=SMALL_FONT,
        command=lambda value: draw_spectrum(
            ax,
            canvas,
            float(value)))
    
    exposure_slider.set(10)
    exposure_slider.pack(fill="x", pady=5)
    
    global acquisition_status_label

    acquisition_status_label = tk.Label(
    exposure_box,
    text="Ready to acquire spectrum.",
    bg=PANEL_BG,
    fg=TEXT_DARK,
    font=SMALL_BOLD,
    justify="left",
    anchor="w")
    acquisition_status_label.pack(fill="x", pady=(6, 2))

    # INSTRUMENT CARD
    instrument_box = make_card(controlframe, "INSTRUMENT")

    spectrum_modebtn = make_button(
        instrument_box,
        f"Mode: {observation_mode}",
        switch_mode)
    spectrum_modebtn.pack(fill="x", pady=4)

    # OBSERVATION CARD/ buttons
    obs_box = make_card(controlframe, "OBSERVATION")

# take reading button
    readingbtn = make_button(
        obs_box,
        "TAKE READING — Save & Measure",
        save_and_measure_spectrum,
        primary=True)
    readingbtn.pack(fill="x", pady=5)

# hubble graph button
    hubblebtn = make_button(
        obs_box,
        "Open Hubble Diagram",
        open_hubble_window)
    hubblebtn.pack(fill="x", pady=5)

# open wedge plot button
    wedgebtn = make_button(
        obs_box,
        "Open Wedge Plot",
        lambda: open_gmm_wedge_plot(auto_update=False))
    wedgebtn.pack(fill="x", pady=5)

# close buttoon
    closebtn = make_button(
        controlframe,
        "Close Acquisition Window",
        spec_window.destroy)
    closebtn.pack(fill="x", side="bottom", padx=10, pady=12)
    
    databasebtn = make_button(
    obs_box,
    "Observation Database",
    open_observation_database)
    databasebtn.pack(fill="x", pady=5)

    draw_spectrum(
        ax,
        canvas,
        exposure_slider.get())
##################################################
# move the telescope faster

moving_north = False
moving_south = False
moving_east = False
moving_west = False


def start_move_north(event):
    global moving_north
    moving_north = True
    move_sound.play(-1) 
    continuous_move_north()


def stop_move_north(event):
    global moving_north
    moving_north = False
    move_sound.stop()


def continuous_move_north():

    if moving_north:

        move_north()

        window.after(25, continuous_move_north)


def start_move_south(event):
    global moving_south
    moving_south = True
    move_sound.play(-1) 
    continuous_move_south()


def stop_move_south(event):
    global moving_south
    moving_south = False
    move_sound.stop()


def continuous_move_south():

    if moving_south:

        move_south()

        window.after(25, continuous_move_south)


def start_move_east(event):
    global moving_east
    moving_east = True
    move_sound.play(-1) 
    continuous_move_east()


def stop_move_east(event):
    global moving_east
    moving_east = False
    move_sound.stop()


def continuous_move_east():

    if moving_east:

        move_east()

        window.after(25, continuous_move_east)


def start_move_west(event):
    global moving_west
    moving_west = True
    move_sound.play(-1) 
    continuous_move_west()


def stop_move_west(event):
    global moving_west
    moving_west = False
    move_sound.stop()


def continuous_move_west():

    if moving_west:

        move_west()

        window.after(25, continuous_move_west)
        
###########################################################
# skymap controls

sky_x_center = 180.0
sky_y_center = 0.0

sky_zoom = 180.0

dragging = False

#########################################################
def open_hubble_window(): #window 2.5

    hubble_window = tk.Toplevel(window)
    hubble_window.title("Hubble Diagram")
    hubble_window.geometry("700x600")

    fig = Figure(figsize=(6,5), dpi=100)

    hubble_ax = fig.add_subplot(111)

    distances = []
    velocities = []

# collect data from galaxy_data
    for _, galaxy in load_spectral_observations().iterrows():
        velocities.append(float(galaxy["Redshift"]) * C_LIGHT)
        distances.append(float(galaxy["Redshift"]) * C_LIGHT / H0)

    # plot points
    hubble_ax.scatter(
        distances,
        velocities,
        color='cyan',
        s=60)

    # labels
    hubble_ax.set_title(
        "Hubble Diagram")

    hubble_ax.set_xlabel(
        "Distance (Mpc)")

    hubble_ax.set_ylabel(
        "Velocity (km/s)")

    hubble_ax.grid(True)

    # optional best-fit line
    if len(distances) > 1 and np.ptp(distances) > 0:

        slope, intercept = np.polyfit(
            distances,
            velocities,
            1)

        xfit = np.linspace(
            min(distances),
            max(distances),
            100)

        yfit = slope * xfit + intercept

        hubble_ax.plot(
            xfit,
            yfit,
            color='red',
            linewidth=2)

    canvas = FigureCanvasTkAgg(
        fig,
        master=hubble_window)

    canvas.draw()

    canvas.get_tk_widget().pack(
        fill="both",
        expand=True)

###########################################################
# build skymap window --> window 0

def build_skymap():

    global root
    global skymap_ax
    global skymap_canvas
    global menuframe
    global current_target

    root = tk.Tk()

    root.title("Interactive Sky Map")

    root.geometry("1000x700")

    navigation = tk.Frame(root, bg=PANEL_BG)
    navigation.pack(fill="x")
    tk.Label(navigation, text="Named landmarks (pointing shortcuts):", bg=PANEL_BG).pack(side="left", padx=8)
    if not STRUCTURES.empty:
        selected_name = tk.StringVar(root, value=str(STRUCTURES.iloc[0]["Name"]))
        tk.OptionMenu(navigation, selected_name, *STRUCTURES["Name"].astype(str)).pack(side="left")
        def go_landmark():
            row = STRUCTURES.loc[STRUCTURES["Name"] == selected_name.get()].iloc[0]
            open_galaxy_finder(dict(Name="All-Sky Survey", RA_deg=float(row.RA_deg), Dec_deg=float(row.Dec_deg)))
        tk.Button(navigation, text="Point Telescope", command=go_landmark).pack(side="left", padx=6)
    def reset_map():
        global sky_x_center, sky_y_center, sky_zoom
        sky_x_center, sky_y_center, sky_zoom = 180.0, 0.0, 180.0
        draw_skymap()
    tk.Button(navigation, text="Show Entire Sky", command=reset_map).pack(side="right", padx=8)

    frame = tk.Frame(root)

    frame.pack(fill="both", expand=True)

    fig = Figure(figsize=(8, 6), dpi=100, facecolor=APP_BG)

    skymap_ax = fig.add_subplot(111)

    skymap_canvas = FigureCanvasTkAgg(
        fig,
        master=frame)

    skymap_canvas.get_tk_widget().pack(
        fill="both",
        expand=True)

    draw_skymap()

    # galaxy clicking
    skymap_canvas.mpl_connect(
    "button_press_event",
    on_skymap_click)

# drag start
    skymap_canvas.mpl_connect(
    "button_press_event",
    start_drag)

# dragging
    skymap_canvas.mpl_connect(
    "motion_notify_event",
    drag_sky)

# drag stop
    skymap_canvas.mpl_connect(
    "button_release_event",
    stop_drag)
    
    skymap_canvas.mpl_connect(
    "scroll_event",
    zoom_skymap)
    
###########################################################
# use mouse to zoom

def zoom_skymap(event):

    global sky_zoom
    global sky_x_center
    global sky_y_center

    if event.xdata is None or event.ydata is None:
        return

    mouse_ra = event.xdata
    mouse_dec = event.ydata

    if event.button == "up":
        scale = 0.8
    elif event.button == "down":
        scale = 1.25
    else:
        return

    sky_x_center = mouse_ra + (sky_x_center - mouse_ra) * scale
    sky_y_center = mouse_dec + (sky_y_center - mouse_dec) * scale
    sky_zoom = np.clip(sky_zoom * scale, 0.5, 180)
    sky_x_center = np.clip(sky_x_center, sky_zoom, 360-sky_zoom)
    sky_y_center = np.clip(sky_y_center, -90+sky_zoom/2, 90-sky_zoom/2)

    draw_skymap()

###########################################################
# map creation: window 0

def draw_skymap():
    skymap_ax.clear()
    skymap_ax.set_facecolor(APP_BG)
    skymap_ax.scatter(REAL_CATALOG.RA_deg, REAL_CATALOG.Dec_deg,
        s=1.5, color="#64748b", alpha=0.45, linewidths=0,
        label="Catalog positions")
    skymap_ax.scatter(OBSERVABLE_CATALOG.RA_deg, OBSERVABLE_CATALOG.Dec_deg,
        s=2, color="#38bdf8", alpha=0.7, linewidths=0,
        label=f"Lab targets: 0 < z ≤ {MAX_LAB_REDSHIFT}")
    if not STRUCTURES.empty:
        scatter = skymap_ax.scatter(STRUCTURES.RA_deg % 360, STRUCTURES.Dec_deg,
            s=45, facecolors="none", edgecolors="#fbbf24", label="Navigation landmarks")
        cursor = mplcursors.cursor(scatter, hover=True)
        @cursor.connect("add")
        def on_hover(sel):
            row = STRUCTURES.iloc[int(sel.index)]
            sel.annotation.set_text(f"{row['Name']}\nNavigation landmark only; targets come from the shared catalog.")
        skymap_canvas._landmark_cursor = cursor
    skymap_ax.set_xlim(max(0, sky_x_center-sky_zoom), min(360, sky_x_center+sky_zoom))
    skymap_ax.set_ylim(max(-90, sky_y_center-sky_zoom/2), min(90, sky_y_center+sky_zoom/2))
    skymap_ax.set_title("All-Sky Galaxy Catalog — Right-click anywhere to point the telescope", color="white")
    skymap_ax.set_xlabel("Right Ascension (degrees)", color="white")
    skymap_ax.set_ylabel("Declination (degrees)", color="white")
    skymap_ax.tick_params(colors="white")
    skymap_ax.grid(alpha=0.15)
    skymap_ax.legend(loc="lower left", facecolor=APP_BG, labelcolor="white", fontsize=8)
    skymap_canvas.draw_idle()


def start_drag(event):

    global dragging
    global last_mouse_x
    global last_mouse_y

    # Only left-click drags
    if event.button != 1:
        return

    if event.xdata is None or event.ydata is None:
        return

    dragging = True
    last_mouse_x = event.xdata
    last_mouse_y = event.ydata

def drag_sky(event):

    global sky_x_center
    global sky_y_center
    global last_mouse_x
    global last_mouse_y

    if not dragging:
        return

    if event.xdata is None or event.ydata is None:
        return

    dx = event.xdata - last_mouse_x
    dy = event.ydata - last_mouse_y

    sky_x_center -= dx
    sky_y_center -= dy
    sky_x_center = np.clip(sky_x_center, sky_zoom, 360-sky_zoom)
    sky_y_center = np.clip(sky_y_center, -90+sky_zoom/2, 90-sky_zoom/2)

    last_mouse_x = event.xdata
    last_mouse_y = event.ydata

    draw_skymap()

# STOP DRAG

def stop_drag(event):

    global dragging

    dragging = False


# allow user to click to choose galaxy
def on_skymap_click(event):
    if event.inaxes != skymap_ax or event.xdata is None or event.ydata is None or event.button not in (2, 3):
        return
    ra, dec = float(event.xdata) % 360, float(event.ydata)
    if not -90 <= dec <= 90: return
    # Snap to a real nearby target within half a degree; otherwise retain the chosen pointing.
    candidates = nearby_galaxies(ra, dec, 0.5)
    if not candidates.empty:
        vectors = unit_vectors(candidates.RA_deg, candidates.Dec_deg)
        row = candidates.iloc[np.argmax(vectors @ unit_vectors([ra], [dec])[0])]
        ra, dec = float(row.RA_deg), float(row.Dec_deg)
    open_galaxy_finder(dict(Name="All-Sky Survey", RA_deg=ra, Dec_deg=dec))


def build_galaxy_finder(window):

    global skyframe
    global menuframe
    global ra_label
    global dec_label
    global readinglbl
    global ra_entry
    global dec_entry

    ##################################################
    # LEFT SIDE = SKY

    skyframe = tk.Frame(window, bg="black")
    skyframe.place(relwidth=0.72, relheight=1)

    ##################################################
    # RIGHT SIDE = MENU

    menuframe = tk.Frame(window, bg="lightgray")
    menuframe.place(relx=0.72, relwidth=0.28, relheight=1)
    
    # telescope control

    telescopelbl = tk.Label(
        menuframe,
        text="Telescope Controls")
        
    telescopelbl.place(
        relx=0.3,
        rely=0.18)

##################################################
# move the telescope

    northbtn = tk.Button(
        menuframe,
        text="N",
        font=("Times New Roman", 14, "bold"),
        width=6,
        height=3)

    northbtn.place(relx=0.37, rely=0.18)
    northbtn.bind("<ButtonPress-1>", start_move_north)
    northbtn.bind("<ButtonRelease-1>", stop_move_north)


    southbtn = tk.Button(
        menuframe,
        text="S",
        font=("Times New Roman", 14, "bold"),
        width=6,
        height=3)

    southbtn.place(relx=0.37, rely=0.35)
    southbtn.bind("<ButtonPress-1>", start_move_south)
    southbtn.bind("<ButtonRelease-1>", stop_move_south)
    
    eastbtn = tk.Button(
        menuframe,
        text="E",
        font=("Times New Roman", 14, "bold"),
        width=6,
        height=3)
    
    eastbtn.place(relx=0.57, rely=0.285)
    eastbtn.bind("<ButtonPress-1>", start_move_east)
    eastbtn.bind("<ButtonRelease-1>", stop_move_east)
    
    westbtn = tk.Button(
        menuframe,
        text="W",
        font=("Times New Roman", 14, "bold"),
        width=6,
        height=3)

    westbtn.place(relx=0.12, rely=0.285)
    westbtn.bind("<ButtonPress-1>", start_move_west)
    westbtn.bind("<ButtonRelease-1>", stop_move_west)
######################################
# telescope alignment
    current_target = None
##################################################
# coordinate labels

    ra_label = tk.Label(
        menuframe,
        text=f"Right Ascension: {ra_center}")
    
    ra_label.place(
        relx=0.10,
        rely=0.60)
    
    dec_label = tk.Label(
        menuframe,
        text=f"Declination: {dec_center}")

    dec_label.place(
        relx=0.10,
        rely=0.65)
    
    ##################################################
# MANUAL COORDINATE ENTRY

    ra_entry = tk.Entry(menuframe)
    ra_entry.place(
        relx=0.10,
        rely=0.73,
        relwidth=0.35)

    ra_entry.insert(0, str(ra_center))

    dec_entry = tk.Entry(menuframe)
    dec_entry.place(
        relx=0.55,
        rely=0.73,
        relwidth=0.35)

    dec_entry.insert(0, str(dec_center))

##################################################
# SET COORDINATES BUTTON

    setcoordbtn = tk.Button(
        menuframe,
        text="Set Coordinates", 
        font=("Times New Roman", 12),
        width=20,
        height=2,
        command=set_coordinates)

    setcoordbtn.place(
        relx=0.1,
        rely=0.83,
        relwidth=0.80)

##################################################
# quit button

    quit_button = tk.Button(
        menuframe,
        text="Quit Simulation",
        command=window.destroy)

    quit_button.place(
        relwidth=0.7,
        relx=0.15,
        rely=0.95)


    ##################################################
    # STATUS LABEL

    readinglbl = tk.Label(
        menuframe,
        text="Galaxy Finder Ready")

    readinglbl.place(relx=0.15, rely=0.05)

    ##################################################
    # COORDINATE LABELS

    ra_label = tk.Label(
        menuframe,
        text=f"Right Ascension: {ra_center}")

    ra_label.place(relx=0.15, rely=0.65)

    dec_label = tk.Label(
        menuframe,
        text=f"Declination: {dec_center}")

    dec_label.place(relx=0.15, rely=0.70)

#############################
def make_button(parent, text, command=None, primary=False):
    return tk.Button(
        parent,
        text=text,
        command=command,
        font=BUTTON_FONT,
        bg=PRIMARY_BUTTON_BG if primary else BUTTON_BG,
        activebackground=PRIMARY_BUTTON_ACTIVE if primary else "#e5e7eb",
        fg=TEXT_DARK,
        relief="ridge",
        bd=2,
        height=2,
        cursor="hand2")


def make_card(parent, title):
    card = tk.LabelFrame(
        parent,
        text=title,
        font=SMALL_BOLD,
        bg=PANEL_BG,
        fg=TEXT_DARK,
        padx=10,
        pady=8,
        bd=1,
        relief="groove")
    card.pack(fill="x", padx=10, pady=6)
    return card


def card_label(parent, left, right="", color=TEXT_DARK):
    row = tk.Frame(parent, bg=PANEL_BG)
    row.pack(fill="x", pady=2)

    tk.Label(
        row,
        text=left,
        font=SMALL_BOLD,
        bg=PANEL_BG,
        fg=TEXT_DARK,
        anchor="w").pack(side="left")

    tk.Label(
        row,
        text=right,
        font=SMALL_BOLD,
        bg=PANEL_BG,
        fg=color,
        anchor="e").pack(side="right")
###########################################################
def load_spectral_observations():

    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame()

    try:
        df = pd.read_excel(
            EXCEL_FILE,
            sheet_name="Spectral Data")

        # Remove accidental blank/header rows
        df = df.dropna(
            subset=["Galaxy", "Redshift"],
            how="any")

        return df

    except Exception as error:
        print("Could not load observation database:", error)
        return pd.DataFrame()
##############################################
def open_observation_database():

    global db_tree
    global db_summary_labels
    global db_window

    # Bring the existing window forward instead of opening duplicates
    try:
        if db_window.winfo_exists():
            refresh_observation_database()
            db_window.lift()
            return
    except:
        pass

    db_window = tk.Toplevel(window)
    db_window.title("LSS Observatory — Observation Database")
    db_window.geometry("1200x750+80+80")
    db_window.minsize(1000, 650)
    db_window.configure(bg=APP_BG)

    # ======================================================
    # MAIN LAYOUT
    # ======================================================

    table_frame = tk.Frame(
        db_window,
        bg=APP_BG,
        highlightbackground="#334155",
        highlightthickness=2)
    table_frame.pack(
        side="left",
        fill="both",
        expand=True,
        padx=(12, 6),
        pady=12)

    control_frame = tk.Frame(
        db_window,
        bg=PANEL_BG,
        width=330)
    control_frame.pack(
        side="right",
        fill="y",
        padx=(6, 12),
        pady=12)
    control_frame.pack_propagate(False)

    # ======================================================
    # TITLE
    # ======================================================

    tk.Label(
        control_frame,
        text="OBSERVATION\nDATABASE",
        font=HEADER_FONT,
        bg=PANEL_BG,
        fg=TEXT_DARK ).pack(pady=(18, 14))

    # ======================================================
    # DARK TREEVIEW STYLE
    # ======================================================

    style = ttk.Style(db_window)

    try:
        style.theme_use("clam")
    except:
        pass

    style.configure(
        "Observation.Treeview",
        background="#0b1623",
        foreground="white",
        fieldbackground="#0b1623",
        rowheight=30,
        font=SMALL_FONT,
        borderwidth=0)

    style.configure(
        "Observation.Treeview.Heading",
        background="#172334",
        foreground="white",
        font=SMALL_BOLD,
        relief="flat")

    style.map(
        "Observation.Treeview",
        background=[("selected", "#166534")],
        foreground=[("selected", "white")])

    # ======================================================
    # TABLE
    # ======================================================

    columns = (
        "Galaxy",
        "Redshift",
        "Velocity",
        "Distance",
        "RA",
        "Dec")

    db_tree = ttk.Treeview(
        table_frame,
        columns=columns,
        show="headings",
        style="Observation.Treeview")

    headings = {
        "Galaxy": "Galaxy ID",
        "Redshift": "Redshift (z)",
        "Velocity": "Velocity (km/s)",
        "Distance": "Distance (Mpc)",
        "RA": "RA (deg)",
        "Dec": "Dec (deg)"}

    widths = {
        "Galaxy": 130,
        "Redshift": 110,
        "Velocity": 150,
        "Distance": 135,
        "RA": 115,
        "Dec": 115}

    for column in columns:
        db_tree.heading(column, text=headings[column])
        db_tree.column(
            column,
            width=widths[column],
            anchor="center")

    y_scrollbar = ttk.Scrollbar(
        table_frame,
        orient="vertical",
        command=db_tree.yview)

    x_scrollbar = ttk.Scrollbar(
        table_frame,
        orient="horizontal",
        command=db_tree.xview)

    db_tree.configure(
        yscrollcommand=y_scrollbar.set,
        xscrollcommand=x_scrollbar.set)

    db_tree.grid(
        row=0,
        column=0,
        sticky="nsew")

    y_scrollbar.grid(
        row=0,
        column=1,
        sticky="ns")

    x_scrollbar.grid(
        row=1,
        column=0,
        sticky="ew")

    table_frame.grid_rowconfigure(0, weight=1)
    table_frame.grid_columnconfigure(0, weight=1)

    # ======================================================
    # SUMMARY
    # ======================================================

    summary_box = make_card(
        control_frame,
        "SUMMARY")

    db_summary_labels = {}

    for key, title in [
        ("total", "Observations:"),
        ("mean_z", "Mean z:"),
        ("mean_distance", "Mean Distance:")]:
        row_frame = tk.Frame(
            summary_box,
            bg=PANEL_BG)
        row_frame.pack(fill="x", pady=4)

        tk.Label(
            row_frame,
            text=title,
            bg=PANEL_BG,
            fg=TEXT_DARK,
            font=SMALL_BOLD
        ).pack(side="left")

        value_label = tk.Label(
            row_frame,
            text="0",
            bg=PANEL_BG,
            fg=TEXT_BLUE,
            font=SMALL_BOLD)
        value_label.pack(side="right")

        db_summary_labels[key] = value_label

    # ======================================================
    # ANALYSIS
    # ======================================================

    analysis_box = make_card(
        control_frame,
        "ANALYSIS")

    make_button(
        analysis_box,
        "Open Wedge Plot",
        lambda: open_gmm_wedge_plot(
            auto_update=False)).pack(fill="x", pady=5)

    make_button(
        analysis_box,
        "Open Hubble Diagram",
        open_hubble_window).pack(fill="x", pady=5)

    make_button(
        analysis_box,
        "Refresh Database",
        refresh_observation_database).pack(fill="x", pady=5)

    make_button(
        control_frame,
        "Close Database",
        db_window.destroy).pack(fill="x",
        side="bottom",
        padx=10,
        pady=12)

    refresh_observation_database()
##############################################
def refresh_observation_database():

    # Make sure the database window/table exists
    if "db_tree" not in globals():
        print("Observation Database is not open.")
        return

    try:
        if not db_tree.winfo_exists():
            return
    except tk.TclError:
        return

    # Remove every currently displayed row
    for item in db_tree.get_children():
        db_tree.delete(item)

    try:
        # Read a completely fresh copy from Excel
        df = pd.read_excel(
            EXCEL_FILE,
            sheet_name="Spectral Data")

        # Remove blank rows
        df = df.dropna(
            subset=["Galaxy", "Redshift"],
            how="any")

    except Exception as error:
        print("DATABASE REFRESH ERROR:", error)

        db_summary_labels["total"].config(text="0")
        db_summary_labels["mean_z"].config(text="0.00000")
        db_summary_labels["mean_distance"].config(text="0.0 Mpc")
        return

    # If Excel now contains no observations, leave the table empty
    if df.empty:
        db_summary_labels["total"].config(text="0")
        db_summary_labels["mean_z"].config(text="0.00000")
        db_summary_labels["mean_distance"].config(text="0.0 Mpc")

        if "acquisition_status_label" in globals():
            try:
                acquisition_status_label.config(
                    text="Database refreshed — no saved observations.",
                    fg=TEXT_BLUE)
            except tk.TclError:
                pass

        print("Database refreshed: 0 observations.")
        return

    inserted_rows = 0

    # Rebuild the visible table from the fresh Excel data
    for _, row in df.iterrows():

        try:
            galaxy = str(row["Galaxy"])
            redshift = float(row["Redshift"])
            velocity = float(row["Velocity (km/s)"])
            distance = float(row["Distance (Mpc)"])
            ra = float(row["RA"])
            dec = float(row["DEC"])

        except (ValueError, TypeError, KeyError):
            continue

        db_tree.insert(
            "",
            "end",
            values=(
                galaxy,
                f"{redshift:.5f}",
                f"{velocity:.0f}",
                f"{distance:.1f}",
                f"{ra:.2f}",
                f"{dec:.2f}"))

        inserted_rows += 1

    # Recalculate summary values from the refreshed data
    valid_z = pd.to_numeric(
        df["Redshift"],
        errors="coerce").dropna()

    valid_distance = pd.to_numeric(
        df["Distance (Mpc)"],
        errors="coerce").dropna()

    db_summary_labels["total"].config(
        text=str(inserted_rows))

    db_summary_labels["mean_z"].config(
        text=(
            f"{valid_z.mean():.5f}"
            if not valid_z.empty
            else "0.00000"))

    db_summary_labels["mean_distance"].config(
        text=(
            f"{valid_distance.mean():.1f} Mpc"
            if not valid_distance.empty
            else "0.0 Mpc"))

    if "acquisition_status_label" in globals():
        try:
            acquisition_status_label.config(
                text=f"Database refreshed — {inserted_rows} observations loaded.",
                fg=TEXT_BLUE)
        except tk.TclError:
            pass

    print(
        f"Database refreshed from Excel: "
        f"{inserted_rows} observations loaded.")
    
##############################################
# OPEN GALAXY FINDER

def open_galaxy_finder(selected_structure):

    global window
    global ra_center
    global dec_center
    global skyframe
    global menuframe
    global ra_label
    global dec_label
    global readinglbl
    global domebtn
    global viewbtn
    global modebtn
    global ra_entry
    global dec_entry
    global acquisitionbtn

    global ACTIVE_STRUCTURE
    global ACTIVE_STRUCTURE_CATALOG
    
    global skyfig
    global sky_ax
    global skyplot
    
    global current_target
    previous = globals().get("window")
    if previous is not None and previous.winfo_exists():
        ra_center = float(selected_structure["RA_deg"]) % 360
        dec_center = float(selected_structure["Dec_deg"])
        ACTIVE_STRUCTURE = selected_structure
        current_target = None
        window.title("LSS OBSERVATORY — All-Sky Survey")
        ra_entry.delete(0, tk.END); ra_entry.insert(0, f"{ra_center:.4f}")
        dec_entry.delete(0, tk.END); dec_entry.insert(0, f"{dec_center:.4f}")
        update_coordinate_labels()
        draw_sky()
        redraw_spectrum_if_open()
        window.lift()
        return
    current_target = None

    ACTIVE_STRUCTURE = selected_structure

    ACTIVE_STRUCTURE_CATALOG = OBSERVABLE_CATALOG

    ra_center = selected_structure["RA_deg"]
    dec_center = selected_structure["Dec_deg"]
    
    ##################################################
# Start a fresh telescope canvas for this structure
##################################################

    skyfig = None
    sky_ax = None
    skyplot = None

# Generate this structure's permanent synthetic sky
    generate_star_field()

    ##################################################
    # MAIN WINDOW

    window = tk.Toplevel(root)

    window.title(f"LSS OBSERVATORY — {selected_structure['Name']}")
    window.geometry("1450x900+40+40")
    window.minsize(1250, 780)
    window.configure(bg=APP_BG)

    # ======================================================
    # MAIN LAYOUT
    # ======================================================

    LEFT_PANEL_WIDTH = 360

    left_container = tk.Frame(
    window,
    bg=PANEL_BG,
    width=LEFT_PANEL_WIDTH)
    
    left_container.pack(side="left", fill="y")
    left_container.pack_propagate(False)

    left_canvas = tk.Canvas(
    left_container,
    bg=PANEL_BG,
    highlightthickness=0)

    left_canvas.pack(side="left", fill="both", expand=True)

    left_scrollbar = tk.Scrollbar(
    left_container,
    orient="vertical",
    command=left_canvas.yview)

    left_scrollbar.pack(side="right", fill="y")

    menuframe = tk.Frame(
    left_canvas,
    bg=PANEL_BG)

    left_canvas.create_window(
    (0, 0),
    window=menuframe,
    anchor="nw",
    width=LEFT_PANEL_WIDTH - 18)

    left_canvas.configure(yscrollcommand=left_scrollbar.set)

    def update_scroll_region(event):
        left_canvas.configure(scrollregion=left_canvas.bbox("all"))

    menuframe.bind("<Configure>", update_scroll_region)
    
    mainframe = tk.Frame(window, bg=APP_BG)
    mainframe.pack(side="right", fill="both", expand=True)

    # ======================================================
    # LEFT CONTROL PANEL
    # ======================================================

    titlelbl = tk.Label(
        menuframe,
        text="⚛  LSS OBSERVATORY",
        font=HEADER_FONT,
        bg=PANEL_BG,
        fg=TEXT_DARK,
        anchor="w")
    titlelbl.pack(fill="x", padx=18, pady=(18, 8))

    target_box = make_card(menuframe, "TARGET REGION")
    tk.Label(
        target_box,
        text=selected_structure["Name"],
        font=TITLE_FONT,
        bg=PANEL_BG,
        fg=TEXT_BLUE,
        anchor="w").pack(fill="x")

    object_box = make_card(menuframe, "CURRENT OBJECT")

    object_status = tk.Label(
        object_box,
        text="● SELECTED" if current_target is not None else "○ NOT CENTERED",
        font=SMALL_BOLD,
        bg=PANEL_BG,
        fg=TEXT_GREEN if current_target is not None else "gray",
        anchor="e")
    object_status.pack(fill="x")

    current_id = "----"
    current_ra = "----"
    current_dec = "----"
    current_z = "----"
    current_v = "----"

    if current_target is not None:
        current_id = str(current_target["Galaxy_ID"])
        current_ra = f"{float(current_target['RA_deg']):.2f}°"
        current_dec = f"{float(current_target['Dec_deg']):.2f}°"
        current_z = "Not measured"
        current_v = "Not measured"

    card_label(object_box, "Galaxy ID:", current_id, TEXT_BLUE)
    card_label(object_box, "RA:", current_ra, TEXT_BLUE)
    card_label(object_box, "Dec:", current_dec, TEXT_BLUE)
    card_label(object_box, "Redshift:", current_z, TEXT_BLUE)
    card_label(object_box, "Velocity:", current_v, TEXT_BLUE)

    coord_box = make_card(menuframe, "COORDINATES")

    ra_label = tk.Label(
        coord_box,
        text=f"RA Center: {ra_center:.2f}°",
        bg=PANEL_BG,
        fg=TEXT_DARK,
        font=SMALL_FONT,
        anchor="w")
    ra_label.pack(fill="x")

    dec_label = tk.Label(
        coord_box,
        text=f"Dec Center: {dec_center:.2f}°",
        bg=PANEL_BG,
        fg=TEXT_DARK,
        font=SMALL_FONT,
        anchor="w")
    dec_label.pack(fill="x", pady=(0, 6))

    ra_entry = tk.Entry(coord_box, font=BUTTON_FONT)
    ra_entry.insert(0, str(round(ra_center, 2)))
    ra_entry.pack(fill="x", pady=3)

    dec_entry = tk.Entry(coord_box, font=BUTTON_FONT)
    dec_entry.insert(0, str(round(dec_center, 2)))
    dec_entry.pack(fill="x", pady=3)

    setcoordbtn = make_button(
        coord_box,
        "⊕  Go To Coordinates",
        set_coordinates)
    setcoordbtn.pack(fill="x", pady=(8, 0))

    telescope_box = make_card(menuframe, "TELESCOPE CONTROLS")

    move_grid = tk.Frame(telescope_box, bg=PANEL_BG)
    move_grid.pack(pady=2)

    northbtn = make_button(move_grid, "N")
    northbtn.grid(row=0, column=1, padx=5, pady=5)

    eastbtn = make_button(move_grid, "E")
    eastbtn.grid(row=1, column=0, padx=5, pady=5)

    center_lbl = tk.Label(
        move_grid,
        text="⊕",
        font=("Times New Roman", 16, "bold"),
        bg=PANEL_BG,
        fg=TEXT_DARK)
    center_lbl.grid(row=1, column=1, padx=8, pady=5)
    
    westbtn = make_button(move_grid, "W")
    westbtn.grid(row=1, column=2, padx=5, pady=5)

    southbtn = make_button(move_grid, "S")
    southbtn.grid(row=2, column=1, padx=5, pady=5)

    northbtn.bind("<ButtonPress-1>", start_move_north)
    northbtn.bind("<ButtonRelease-1>", stop_move_north)

    southbtn.bind("<ButtonPress-1>", start_move_south)
    southbtn.bind("<ButtonRelease-1>", stop_move_south)

    eastbtn.bind("<ButtonPress-1>", start_move_east)
    eastbtn.bind("<ButtonRelease-1>", stop_move_east)

    westbtn.bind("<ButtonPress-1>", start_move_west)
    westbtn.bind("<ButtonRelease-1>", stop_move_west)

    instrument_box = make_card(menuframe, "INSTRUMENT")

# dome button
    domebtn = make_button(
        instrument_box,
        "Dome: Open" if dome_open else "Dome: Closed",
        toggle_dome)
    domebtn.pack(fill="x", pady=3)
# view button
    viewbtn = make_button(
        instrument_box,
        f"View: {scope_mode}",
        switch_view)
    viewbtn.pack(fill="x", pady=3)
# mode button
    modebtn = make_button(
        instrument_box,
        f"Mode: {observation_mode}",
        switch_mode)
    modebtn.pack(fill="x", pady=3)

    obs_box = make_card(menuframe, "OBSERVATION")
# take reading button
    readingbtn = make_button(
        obs_box,
        "◎  TAKE READING",
        take_reading,
        primary=True)
    readingbtn.pack(fill="x", pady=4)
# MOS window button
    multi_object_btn = make_button(
        obs_box,
        "Multi-Object Spectroscopy",
        open_multi_object_window)
    multi_object_btn.pack(fill="x", pady=4)
# quit button
    quit_button = make_button(
        menuframe,
        "Quit Simulation",
        window.destroy)
    quit_button.pack(fill="x", side="bottom", padx=10, pady=12)

    # ======================================================
    # RIGHT OBSERVATORY PANEL
    # ======================================================

    topbar = tk.Frame(
        mainframe,
        bg="#0b1623",
        height=72,
        highlightbackground="#334155",
        highlightthickness=1)
    topbar.pack(side="top", fill="x", padx=12, pady=(12, 6))
    topbar.pack_propagate(False)

    tk.Label(
        topbar,
        text=f"RA (Center):  {ra_center:.2f}°\nDec (Center): {dec_center:.2f}°",
        font=BUTTON_FONT,
        bg="#0b1623",
        fg="white",
        justify="left").pack(side="left", padx=25)

    tk.Label(
        topbar,
        text=(
            f"Field of View: "
            f"{get_current_fov():.1f}° × "
            f"{get_current_fov():.1f}°\n"
            f"Detector: Synthetic CCD"
        ),
        font=BUTTON_FONT,
        bg="#0b1623",
        fg="white",
        justify="left"
        ).pack(side="left", padx=35)

    tk.Label(
        topbar,
        text=f"View: {scope_mode}\nMode: {observation_mode}",
        font=BUTTON_FONT,
        bg="#0b1623",
        fg="white",
        justify="left").pack(side="left", padx=35)

    global dome_status_label

    dome_status_label = tk.Label(
    topbar,
    text=f"Dome: {'Open' if dome_open else 'Closed'}",
    font=BUTTON_FONT,
    bg="#0b1623",
    fg="#22c55e" if dome_open else "#f87171",
    justify="left")
    
    dome_status_label.pack(side="left", padx=35)

    tk.Label(
        topbar,
        text="●  Ready",
        font=TITLE_FONT,
        bg="#0b1623",
        fg="#22c55e",
        padx=20,
        pady=8,
        relief="ridge",
        bd=1).pack(side="right", padx=20)

    skyframe = tk.Frame(
        mainframe,
        bg="black",
        highlightbackground="#334155",
        highlightthickness=1)
    skyframe.pack(side="top", fill="both", expand=True, padx=12, pady=6)

    bottombar = tk.Frame(
        mainframe,
        bg="#0b1623",
        height=42,
        highlightbackground="#334155",
        highlightthickness=1)
    bottombar.pack(side="bottom", fill="x", padx=12, pady=(6, 12))
    bottombar.pack_propagate(False)

    tk.Label(
        bottombar,
        text=f"All-sky lab targets: {len(OBSERVABLE_CATALOG):,}",
        font=BUTTON_FONT,
        bg="#0b1623",
        fg="white").pack(side="left", padx=20)

    selected_text = "Selected Object: None"
    if current_target is not None:
        selected_text = f"Selected Object: Galaxy ID {current_target['Galaxy_ID']}"

    tk.Label(
        bottombar,
        text=selected_text,
        font=BUTTON_FONT,
        bg="#0b1623",
        fg="white").pack(side="left", padx=80)

    tk.Label(
        bottombar,
        text="Exposure: 10 s",
        font=BUTTON_FONT,
        bg="#0b1623",
        fg="white").pack(side="right", padx=25)

##################################################
# INITIAL TELESCOPE DRAW
##################################################

    draw_sky()

    #########################################
def clear_all_data():

    global galaxy_data

    if not messagebox.askyesno("Clear observations", "Delete all saved observations? This cannot be undone."):
        return

    try:
        workbook = openpyxl.load_workbook(EXCEL_FILE)

        for sheet_name in [
            "Spectral Data",
            "Photometry Data",
            "Wedge Map Data"]:
            if sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]

                # Delete every row except the header row
                if sheet.max_row >= 2:
                    sheet.delete_rows(
                        idx=2,
                        amount=sheet.max_row - 1)
                    

        workbook.save(EXCEL_FILE)
        workbook.close()

        galaxy_data.clear()
        print("Excel database cleared.")

    except Exception as error:
        print("RESET ERROR:", error)
        return

    # Empty the visible database immediately
    if "db_tree" in globals():
        try:
            for item in db_tree.get_children():
                db_tree.delete(item)

            db_summary_labels["total"].config(text="0")
            db_summary_labels["mean_z"].config(text="0.00000")
            db_summary_labels["mean_distance"].config(text="0.0 Mpc")

        except (tk.TclError, NameError):
            pass

    if "wedge_ax" in globals() and "wedge_canvas" in globals():
        try:
            update_wedge_plot()
        except Exception:
            pass

    if "acquisition_status_label" in globals():
        try:
            acquisition_status_label.config(
                text="All saved data cleared.",
                fg="#15803d")
        except tk.TclError:
            pass
    
build_skymap()
root.mainloop()
