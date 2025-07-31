# FlotaMaster Analyzer

FlotaMaster Analyzer is an Excel 2016+ VBA tool that automates end-to-end analysis of metallurgical flotation test data. With one click, it computes weighted copper performance scores, assesses impurities and kinetics, identifies the best collector, generates formatted expert comments and optimization suggestions, and visualizes results on a dashboard.

---

## Table of Contents

1. [Overview](#overview)  
2. [Features](#features)  
3. [Installation](#installation)  
4. [Usage](#usage)  
5. [Project Structure & Components](#project-structure--components)  
6. [Dependencies](#dependencies)  
7. [Developer Notes](#developer-notes)  

---

## Overview

FlotaMaster Analyzer guides metallurgical engineers through rapid, repeatable flotation test evaluations:

- Computes weighted Cu recovery & grade scores  
- Flags pyrite/carbon impurities  
- Analyzes flotation kinetics (rate vs. time)  
- Ranks collectors (C14, C26, C38, C50)  
- Generates a formatted expert report (comments + suggestions)  
- Builds performance & kinetics charts on a dashboard sheet  
- Digitally signs the VBA project for integrity  

---

## Features

- Compatibility with Excel 2016+ (Windows)  
- One-click execution via custom ribbon button  
- Interactive Settings Pane for weights, thresholds, comment templates  
- Automated input validation, error logging & notifications  
- Modular, maintainable VBA codebase  
- Automated dashboard chart creation  
- Configurable report templates & exportable output  
- Sub-5-second runtime for up to 4 collectors ? 4 timepoints  
- VBA project digital signature for security  

---

## Installation

1. Ensure you have **Microsoft Excel 2016 or later** (Windows) installed.  
2. Download or clone the repository to your local machine.  
3. Unblock the ZIP (if downloaded) and extract all files into a folder.  
4. Open `flotamaster analyzer.xlsm` in Excel.  
5. Go to **File ? Options ? Trust Center ? Trust Center Settings?**  
   - Enable macros under **Macro Settings**.  
   - (Optional) Add the folder to **Trusted Locations**.  
6. If prompted, enable the VBA digital signature.  

---

## Usage

1. Open **flotamaster analyzer.xlsm**.  
2. The custom ribbon and Settings Pane load automatically.  
3. In the **Settings Pane**, configure:  
   - Weight constants for recovery vs. grade  
   - Impurity & kinetics thresholds  
   - Comment & suggestion templates  
4. On the ?Input? sheet, enter:  
   - Collector names (cells V18:V21, V30:V33, V42:V45, V54:V57)  
   - Mass pull data (same ranges)  
   - Cu grade data (cells W18:W21, W30:W33, W42:W45, W54:W57)  
   - Pyrite data (cells AH18:AH21, AH30:AH33, AH42:AH45, AH54:AH57)  
5. Click **Analyze** on the custom ribbon.  
6. Review:  
   - **Dashboard** sheet (performance & kinetics charts starting at E5)  
   - **Report** section (bullet-point analysis + optimization suggestions)  
   - **Log** sheet (hidden; execution & error details)  

---

## Project Structure & Components

All VBA modules and helper files are embedded in `flotamaster analyzer.xlsm`. The repository also includes standalone `.bas`/`.cls`/`.xml` files for version control.

### Core VBA Modules

- **configmodule.bas**  
  ? Status: Pass  
  ? Purpose: Load/save configurable weights, thresholds, and comment templates.  
  ? Dependencies: ?  

- **utils.bas**  
  ? Status: Pass  
  ? Purpose: Safe parsing, formatting, logging helper routines.  
  ? Dependencies: ?  

- **inputhandler.bas**  
  ? Status: Pass  
  ? Purpose: Read & validate user inputs from predefined ranges.  
  ? Dependencies: Utils  

- **scoringengine.bas**  
  ? Status: Pass  
  ? Purpose: Compute & normalize weighted Cu recovery/grade scores.  
  ? Dependencies: ConfigModule  

- **impuritymodule.bas**  
  ? Status: Pass  
  ? Purpose: Evaluate pyrite/carbon impurity levels vs. thresholds.  
  ? Dependencies: ConfigModule  

- **kineticsmodule.bas**  
  ? Status: Pass  
  ? Purpose: Analyze flotation kinetics time-series & rate trends.  
  ? Dependencies: ConfigModule  

- **chartmodule.bas**  
  ? Status: Pass  
  ? Purpose: Generate & format dashboard performance/kinetics charts.  
  ? Dependencies: Utils  

- **reportgenerator.bas**  
  ? Status: Pass  
  ? Purpose: Clear prior output; write bullet-point analysis & suggestions.  
  ? Dependencies: InputHandler  

- **loggingmodule.bas**  
  ? Status: Pass  
  ? Purpose: Initialize logger; record info, warnings, errors to hidden sheet.  
  ? Dependencies: Utils  

- **settingsmodule.bas**  
  ? Status: Pass  
  ? Purpose: Manage interactive Settings Pane (load/save settings).  
  ? Dependencies: CustomUI  

- **digitalsignature.bas**  
  ? Status: Pass  
  ? Purpose: Verify & apply digital signature to the VBA project.  
  ? Dependencies: ?  

### Project & UI Files

- **thisworkbook.cls**  
  ? Status: Updated  
  ? Purpose: Workbook event handling (? load ribbon, initialize pane).  
  ? Dependencies: CustomUI  

- **customui.xml**  
  ? Status: Pass  
  ? Purpose: Defines custom ribbon buttons & callback mappings.  
  ? Dependencies: ?  

- **flotamaster analyzer.xlsm**  
  ? Status: Pass  
  ? Purpose: Main Excel workbook (embedded modules, input/output sheets).  
  ? Dependencies: ThisWorkbook, all modules  

- **analyzer.xlsm**  
  ? Status: Fail (deprecated)  
  ? Purpose: Legacy/failed artifact; not used in current workflow.  

---

## Dependencies

- Microsoft Excel 2016 or later (Windows)  
- No external libraries or COM add-ins  
- VBA digital signature (self-signed or corporate)  

---

## Developer Notes

- All modules follow a **modular architecture**; new analysis routines can be added as separate `.bas` files.  
- **Pseudo-code** for each module lives in its header comments; see the top of each `.bas` for function outlines.  
- To **rebuild** the ribbon XML, edit `customui.xml` and reload the workbook.  
- For **debugging**, open the VBA editor (Alt + F11) and set breakpoints or step through code.  
- Logs are written to a hidden sheet named `Log`; call `FlushLogs` to force a write.  

---

Thank you for using FlotaMaster Analyzer! If you encounter issues or have feature requests, please open an issue in the repository.