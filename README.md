# PyCLEA — Interactive Galaxy Redshift Laboratory

**A Python desktop application for exploring how astronomers use galaxy spectra and redshift measurements to map the large-scale structure of the universe.**

Developed by **Sara Tamura** as part of the **X-SIG summer research program at Gettysburg College (2026)**, PyCLEA adapts the educational goals of the Project CLEA *Large Scale Structure of the Universe* laboratory into an interactive Python application.

PyCLEA combines an astronomical catalog with simulated telescope imaging, spectroscopy, measurement tools, and three-dimensional visualization. Students can select galaxies, acquire synthetic spectra, identify absorption lines, calculate redshifts, and build a map from their own measurements.

> **Portfolio note:** This repository contains a portable demonstration edition of PyCLEA. The included galaxy catalog is a 30,000-row random sample of the research catalog, not a complete or statistically representative map of cosmic structure. Galaxy sky positions and catalog redshifts are survey-derived; detector imagery, noise, and acquired spectra are simulated for educational use.

## Project highlights

- **Interactive telescope interface:** Navigate the sky using right ascension and declination, switch between finder and main-scope views, and operate a simulated observatory dome.
- **Catalog-based galaxy targeting:** Select objects at sky coordinates drawn from a galaxy catalog, with a synthetic CCD-like display and simulated foreground stars.
- **Spectral acquisition:** Generate noisy optical spectra containing redshifted calcium H and K absorption features; adjust exposure time to investigate measurement quality.
- **Hands-on redshift measurement:** Select observed spectral-line wavelengths and calculate galaxy redshift from the measured displacement of the calcium lines.
- **Distance estimation:** Use the low-redshift approximation to estimate recession velocity and Hubble-law distance.
- **3D visualization:** Display student-measured galaxies in an interactive Plotly spatial map.
- **Observation records:** Save spectral and photometric measurements in a locally generated Excel workbook.
- **Additional observing mode:** Explore a simulated multi-object spectroscopy workflow.

## Research poster and project visuals

**[View the X-SIG research poster (PDF)](docs/XSIG_Research_Poster.pdf)** — *PyCLEA: Modernizing the Large Scale Structure of the Universe Lab*, **R. Johnson and Sara Tamura**, Gettysburg College.

The poster presents the project's educational motivation, methods, and representative visualizations, including the **Galaxy Finder interface**, **simulated spectral acquisition**, **all-sky galaxy catalog**, **three-dimensional galaxy redshift map**, and **survey sky coverage**. The plotted survey catalog positions in the poster are distinct from the smaller demonstration dataset bundled with this repository.

The poster serves as the project's visual overview. Separate screenshots of the application may be added later, but are not necessary to browse the project.

## Scientific background

Light emitted or absorbed at a known rest wavelength is observed at a longer wavelength when a galaxy is receding. The redshift is

\[
z = \frac{\lambda_{\mathrm{observed}}-\lambda_{\mathrm{rest}}}{\lambda_{\mathrm{rest}}}.
\]

In PyCLEA, students identify the calcium K and H absorption features and estimate redshift from their measured wavelengths. For the nearby-galaxy teaching exercise, the program uses

\[
v \approx cz, \qquad d \approx \frac{v}{H_0},
\]

with \(H_0 = 70\ \mathrm{km\,s^{-1}\,Mpc^{-1}}\). These are **low-redshift approximations**, not a full cosmological distance calculation.

The program converts each measured distance and sky position into Cartesian coordinates for visualization:

\[
x=d\cos(\delta)\cos(\alpha),\quad
 y=d\cos(\delta)\sin(\alpha),\quad
 z=d\sin(\delta),
\]

where \(\alpha\) is right ascension and \(\delta\) is declination.

## Technologies

**Python**, Tkinter, NumPy, pandas, Matplotlib, SciPy, Plotly, openpyxl, mplcursors, and optional pygame audio support.

The application uses a spatial nearest-neighbor search to identify nearby catalog objects, a synthetic detector model to render telescope views, and a browser-based interactive 3D plot for recorded measurements.

## Installation and running

**Requirements:** Python 3.10 or newer, a desktop environment with Tkinter support, and a web browser for interactive Plotly output.

1. Download or clone this repository and open a terminal in its directory.
2. Install dependencies:

   ```bash
   python -m pip install -r requirements.txt
   ```

3. Launch the application:

   ```bash
   python PyCLEA.py
   ```

   On some macOS installations, the command is `python3` rather than `python`. Tkinter must be available in the Python installation.

4. Use the application to choose a field, target a galaxy, acquire a spectrum, measure the calcium lines, and record an observation. Open the 3D plot after recording measurements.

The program creates `galaxy_database.xlsx` automatically as observations are saved. The 3D map is exported locally as `student_3d_wedge_plot.html`.

## Repository structure

```text
PyCLEA/
├── PyCLEA.py                    # Main desktop application
├── README.md                    # Project overview and setup
├── requirements.txt             # Python dependencies
├── .gitignore                   # Excludes generated/local files
├── data/
│   └── sample_galaxy_catalog.csv
├── assets/
│   └── telescope.wav
├── docs/
│   └── XSIG_Research_Poster.pdf # Research poster and visual overview
├── screenshots/                 # Optional future application screenshots
└── PORTFOLIO_CHECKLIST.md       # Pre-publication checks
```

### Data and reproducibility

The included `data/sample_galaxy_catalog.csv` is a fixed-seed random selection of **30,000 rows** from the larger research catalog. It supports demonstration of the software but does **not** preserve the completeness or selection function of the original survey compilation. It should not be used to draw conclusions about the morphology of the cosmic web.

Required fields are `Galaxy_ID`, `RA_deg`, `Dec_deg`, and `Redshift`. The optional `Survey` column identifies the source when available.

To use a larger catalog, place an appropriately formatted file at `data/clean_master_galaxy_catalog.csv`. The application will load that file in preference to the sample. This full-catalog filename is excluded by `.gitignore` to avoid accidental upload.

An optional `data/known_large_scale_structures.csv` provides named observing targets. The application can run without it.

**Data provenance and redistribution:** Before publishing the sample data publicly, document the original survey sources and verify that redistribution is permitted. The catalog is survey-derived; it is not presented as a new survey produced by this project.

## Scope and limitations

- The simulated telescope image and absorption spectra are educational models, not recorded telescope observations.
- The example data is a reduced demonstration sample and may leave some sky regions sparsely populated.
- The Hubble-law distance approximation is intended for the nearby-galaxy teaching exercise.
- The interactive 3D plot shows **student-recorded observations**. It does not perform DBSCAN or Gaussian-mixture clustering in this version of the application. Machine-learning analysis of cosmic structure is a separate research activity.
- The application has been run successfully on the developer's computer; cross-platform installation and the complete observing workflow have not yet been independently tested.

## Research context and acknowledgments

This project was developed during the **2026 Gettysburg College X-SIG summer research program** and builds on the instructional concepts of the **Project CLEA Large Scale Structure of the Universe** laboratory.

**Poster coauthor:** R. Johnson. **Before public release:** Confirm the preferred full name and acknowledgment of the research advisor and collaborators, obtain approval to share the coauthored poster, verify exact galaxy survey citations and usage terms, and confirm attribution or permission details for the telescope sound effect. A software license has not yet been selected pending confirmation of ownership and publication permissions.

## Author

**Sara Tamura**  
B.S. Physics candidate, Data Science minor — Gettysburg College, expected May 2027  
