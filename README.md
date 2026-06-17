# DrawioGen

A Python CLI tool that generates draw.io-compatible network topology diagrams (`.drawio` XML) from a structured Excel input file. Designed specifically for Metro/IPRAN/Router network environments — supporting multi-vendor node types, automatic role detection from hostname patterns, and multiple auto-layout algorithms.

No manual drawing required. Feed it your NE inventory and link table, and get a ready-to-open diagram in draw.io.

---

## Table of Contents

- [Project Overview](#project-overview)
- [Architecture](#architecture)
- [File Structure](#file-structure)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage Guide](#usage-guide)
  - [Interactive Mode](#interactive-mode)
  - [CLI Mode](#cli-mode)
  - [Layout Selection](#layout-selection)
  - [Listing All Roles](#listing-all-roles)
- [Excel Input Schema](#excel-input-schema)
  - [Nodes Sheet](#nodes-sheet)
  - [Links Sheet](#links-sheet)
- [Node Types Reference](#node-types-reference)
  - [Explicit Node Types](#explicit-node-types-node_type-column)
  - [Auto-Detected Roles](#auto-detected-roles-hostname-pattern-matching)
- [Known Limitations](#known-limitations)
- [Changelog](#changelog)

---

## Project Overview

DrawioGen solves a common pain point in network operations: manually drawing topology diagrams every time a network expands. Instead of dragging and dropping icons in draw.io by hand, you maintain a simple Excel sheet and let the tool generate the diagram for you.

**Key capabilities:**

- Reads node inventory (hostname, loopback IP, node type, site ID) and link table from a single `.xlsx` file
- Outputs a fully structured `.drawio` XML file, ready to open and further edit in draw.io or `app.diagrams.net`
- Supports **5 explicit node types** with pre-assigned colors and Cisco shape icons: `ME`, `BNG`, `CSR`, `OLT`, `BTS`
- Falls back to **automatic role detection** from hostname patterns for unlisted node types (supports 5G NFs, PE/P/Core routers, switches, firewalls, CPEs, and more)
- Offers **5 layout algorithms** for different topology shapes (balanced, hierarchical, force-directed, spectral, flat grid)
- Handles disconnected graph components by tiling them side by side
- Warns on duplicate node entries and unknown link endpoints without crashing
- Runs in **interactive console mode** (prompts for input, output, layout) or **CLI mode** (flags for automation/scripting)

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                     drawio_topology.py                          │
│                                                                 │
│  Input layer                                                    │
│  ├─ read_excel()         Reads Nodes + Links sheets from .xlsx  │
│  ├─ parse_nodes_text()   [Legacy] Reads nodes from .txt file    │
│  └─ parse_links_text()   [Legacy] Reads links from .txt file    │
│                                                                 │
│  Role resolution                                                │
│  ├─ resolve_role()       node_type column → NODE_TYPES dict     │
│  │                       blank → ROLES pattern match on hostname│
│  └─ make_label()         Builds HTML label (2-line or 3-line)   │
│                                                                 │
│  Layout engine                                                  │
│  ├─ layout_nodes()       Dispatcher — selects algorithm         │
│  ├─ kamada_kawai         Balanced (default, via networkx)       │
│  ├─ spring               Force-directed (via networkx)          │
│  ├─ spectral             Eigenvector-based (via networkx)       │
│  ├─ hierarchical         Top-down tree (built-in, no networkx)  │
│  └─ flat                 Role-grouped grid (built-in)           │
│                                                                 │
│  XML builder                                                    │
│  └─ build_xml()          Assembles mxGraphModel XML, calls      │
│                          minidom for pretty-print output         │
│                                                                 │
│  Entry points                                                   │
│  ├─ main()               CLI mode (sys.argv flags)              │
│  └─ run_interactive()    Interactive console session            │
└─────────────────────────────────────────────────────────────────┘
         │                                        │
         ▼ reads                                  ▼ writes
┌─────────────────┐                   ┌────────────────────────┐
│  TopoExcel/     │                   │  OutputTopo/           │
│  topology.xlsx  │                   │  topology.drawio       │
│  (Nodes sheet)  │                   │  (draw.io XML)         │
│  (Links sheet)  │                   └────────────────────────┘
└─────────────────┘
```

**Processing flow:**

1. `read_excel()` reads the `Nodes` and `Links` sheets, normalises column headers, skips blank rows, and warns on duplicates.
2. Each node is passed to `resolve_role()` — explicit `node_type` values take priority over hostname pattern matching.
3. `layout_nodes()` runs the selected algorithm to assign `(x, y)` canvas coordinates to each node.
4. `build_xml()` emits an `mxGraphModel` XML document with a cell for each node (with Cisco shape style strings) and a straight-line edge for each link.
5. The output file is written to `OutputTopo/`.

---

## File Structure

```
DrawioGen/
├── drawio_topology.py     Main script — all logic in a single file
├── TopoExcel/             Place your input .xlsx files here
│   └── (template.xlsx)    Template Excel file included in repo
├── OutputTopo/            Generated .drawio files written here
│                          (created automatically if absent)
└── README.md              This file
```

---

## Prerequisites

- **Python 3.10+** (uses `list[Node]` type hints — will not run on 3.9 or below)
- **pandas** + **openpyxl** — for reading Excel input
- **networkx** *(optional but recommended)* — required for `kamada_kawai`, `spring`, and `spectral` layouts; falls back to `hierarchical` if not installed

All other dependencies (`xml.etree`, `argparse`, `dataclasses`, `re`, `os`, `sys`) are Python standard library.

---

## Installation

**1. Clone the repository**

```bash
git clone https://github.com/Flalemuel/DrawioGen.git
cd DrawioGen
```

**2. Install Python dependencies**

```bash
pip install pandas openpyxl networkx
```

> If you only plan to use `--layout hierarchical` or `--layout flat`, you can skip `networkx`.

**3. Verify**

```bash
python drawio_topology.py --list-roles
```

This prints all supported explicit node types and auto-detection patterns without requiring any input file.

---

## Usage Guide

### Interactive Mode

Run the script with no arguments to enter the interactive console. You will be prompted step by step:

```bash
python drawio_topology.py
```

```
 Enter Excel input filename (e.g. topology.xlsx): my_network.xlsx
 Enter output filename (without extension, e.g. my_topology): klaten_topo
 Select layout algorithm:
   1. kamada_kawai     Default — balanced, best for most topologies  <-- default
   2. spring           Dense or mesh networks, organic clustering <-- Best layout
   3. spectral         Very large graphs (100 nodes), fast
   4. hierarchical     Strict top-down trees, no networkx needed
   5. flat             No links, grouped by device type
```

- The input file must be placed inside the `TopoExcel/` folder.
- The output file is written to `OutputTopo/`.
- After each run, you are asked whether to generate another topology — allowing batch work without restarting.

---

### CLI Mode

Pass arguments directly for use in scripts or automated pipelines:

```bash
# Basic usage — default layout (kamada_kawai)
python drawio_topology.py -i TopoExcel/topology.xlsx -o OutputTopo/out.drawio

# Explicit layout selection
python drawio_topology.py -i topology.xlsx -o out.drawio --layout spring
python drawio_topology.py -i topology.xlsx -o out.drawio --layout hierarchical
python drawio_topology.py -i topology.xlsx -o out.drawio --layout spectral
python drawio_topology.py -i topology.xlsx -o out.drawio --layout flat

# Custom sheet names (if your Excel uses non-default names)
python drawio_topology.py -i topology.xlsx -o out.drawio --sheet-nodes NE_Data --sheet-links Link_Data

# Legacy text file input (no Excel)
python drawio_topology.py -n nodes.txt -l links.txt -o out.drawio
```

**All CLI flags:**

| Flag | Short | Default | Description |
|---|---|---|---|
| `--input` | `-i` | — | Excel input file path (`.xlsx`) |
| `--output` | `-o` | `topology.drawio` | Output file path (`.drawio`) |
| `--sheet-nodes` | — | `Nodes` | Name of the Nodes sheet in Excel |
| `--sheet-links` | — | `Links` | Name of the Links sheet in Excel |
| `--layout` | — | `kamada_kawai` | Layout algorithm (see below) |
| `--nodes` | `-n` | — | [Legacy] Nodes text file |
| `--links` | `-l` | — | [Legacy] Links text file |
| `--list-roles` | — | — | Print all node types and patterns, then exit |

---

### Layout Selection

| Algorithm | When to use |
|---|---|
| `kamada_kawai` | **Default.** Balanced, minimises edge length variance. Reads like a real network diagram. Best starting point for most topologies. |
| `spring` | Dense or mesh networks. Organic clustering — lets strongly connected nodes group naturally. Recommended for complex topologies with many cross-links. |
| `spectral` | Very large graphs (100+ nodes). Eigenvector-based placement — fast but less aesthetically balanced than `kamada_kawai`. |
| `hierarchical` | Strict top-down trees. Useful for clean hierarchy diagrams (ME → CSR → OLT). Does not require `networkx`. |
| `flat` | No links provided, or you just want nodes grouped by device type in a grid. Does not require `networkx`. |

> **Disconnected graphs:** For `kamada_kawai` and `spectral`, the tool automatically detects disconnected components and tiles them side by side rather than overlapping them.

---

### Listing All Roles

```bash
python drawio_topology.py --list-roles
```

Prints the full table of explicit `node_type` values and all fallback hostname patterns, then exits. Useful for verifying what a given hostname will resolve to before generating a diagram.

---

## Excel Input Schema

The input file must have **two sheets**: one for nodes and one for links. Default sheet names are `Nodes` and `Links` (case-sensitive). Both can be overridden with `--sheet-nodes` and `--sheet-links`.

A template file is included in the `TopoExcel/` directory.

---

### Nodes Sheet

Column order matters — headers are read positionally (column 1, 2, 3, 4), not by name. The tool normalises and ignores blank rows automatically.

| Column | Position | Required | Description |
|---|---|---|---|
| `hostname / tower_id` | 1 | Yes | Device hostname for routers; tower ID for BTS nodes |
| `ip_loopback` | 2 | No | Loopback IP address (displayed in label below hostname) |
| `node_type` | 3 | No | Explicit device type — see [Node Types Reference](#node-types-reference). Leave blank for auto-detection. |
| `site_id` | 4 | No | Site identifier — used as a third label line for BTS nodes only |

**Example:**

| hostname / tower_id | ip_loopback | node_type | site_id |
|---|---|---|---|
| JKT-ME-01 | 10.1.0.1 | ME | |
| JKT-CSR-01 | 10.1.1.1 | CSR | |
| JKT-OLT-01 | 10.1.2.1 | OLT | |
| TWR-JKT-001 | 10.1.3.1 | BTS | SITE-JKT-001 |

> **Duplicates:** If the same hostname appears more than once, only the first occurrence is used. A warning is printed for each duplicate skipped.

---

### Links Sheet

Defines point-to-point connections between nodes. Each row is one link drawn as a straight line in the diagram.

| Column | Position | Required | Description |
|---|---|---|---|
| `source_hostname` | 1 | Yes | Must exactly match a hostname in the Nodes sheet |
| `destination_hostname` | 2 | Yes | Must exactly match a hostname in the Nodes sheet |

**Example:**

| source_hostname | destination_hostname |
|---|---|
| JKT-ME-01 | JKT-CSR-01 |
| JKT-CSR-01 | JKT-OLT-01 |

> **Unknown endpoints:** If a link references a hostname that does not exist in the Nodes sheet, that link is skipped with a `[warn]` message. Generation continues for all valid links. The final output shows how many links were skipped.

> **Duplicate links:** If the same source→destination pair appears more than once, only one edge is drawn.

> **Links sheet absent:** If no `Links` sheet exists (or the `--sheet-links` name does not match), the tool generates a nodes-only diagram using the `flat` layout automatically.

---

## Node Types Reference

### Explicit Node Types (`node_type` column)

Set the `node_type` column in the Nodes sheet to one of these exact values (case-insensitive):

| Type | Label | Fill Color | Shape Icon | Label format |
|---|---|---|---|---|
| `ME` | Metro Router | Red `#E51400` | Cisco Router | hostname / ip |
| `BNG` | BNG/BRAS | Dark Blue `#5552FF` | Cisco Router | hostname / ip |
| `CSR` | CSR | Cisco Blue `#036897` | Cisco Router | hostname / ip |
| `OLT` | OLT/ONU | Light Pink `#F8CECC` | Cisco Optical Transport | hostname / ip |
| `BTS` | BTS | Cisco Blue `#036897` | Cisco Radio Tower | **tower_id (bold) / ip / site_id** |

> All Cisco shape icons are rendered using the built-in `mxgraph.cisco.*` shape library in draw.io. No external shape packs need to be installed.

---

### Auto-Detected Roles (hostname pattern matching)

When `node_type` is left blank, the tool matches the hostname (lowercased) against keyword patterns. The first match wins.

| Role | Detected patterns | Label | Shape |
|---|---|---|---|
| Core Router | `core`, `backbone`, `bb-` | Core Router | Cisco Router |
| PE Router | `pe-`, `pe_`, `asbr`, `edge-rtr` | PE Router | Cisco Router |
| P Router | `-p-`, `_p_`, `transit-` | P Router | Cisco Router |
| BNG/BRAS | `bng`, `bras` | BNG/BRAS | Cisco Router |
| CSR | `csr`, `asr` | CSR | Cisco Router |
| Metro Router | `-me-`, `_me_`, `me1/2/3` | Metro Router | Cisco Router |
| gNB | `gnb`, `enb`, `ran-`, `radio-`, `site-` | gNB | Cisco Wireless AP |
| BTS | `bts`, `twr-`, `tower-` | BTS | Cisco Radio Tower |
| OLT/ONU | `olt`, `onu`, `ont`, `-fh`, `_fh` | OLT/ONU | Rectangle |
| Switch | `sw-`, `sw_`, `-sw`, `switch`, `agg-`, `nx` | Switch | Cisco Workgroup Switch |
| Firewall | `fw`, `firewall`, `asa-`, `palo-` | Firewall | Cisco Firewall |
| UPF | `upf` | UPF | Cisco Server |
| AMF | `amf` | AMF | Cisco Server |
| SMF | `smf` | SMF | Cisco Server |
| PCF | `pcf` | PCF | Cisco Server |
| 5GC NF | `udm`, `udr`, `ausf`, `nrf`, `nssf`, `nef` | 5GC NF | Cisco Server |
| CPE | `cpe`, `home-` | CPE | Cisco Router |
| Server | `srv`, `server`, `vm-`, `dc-` | Server | Cisco Server |
| *(default)* | *(no match)* | Device | Cisco Router |

> Run `python drawio_topology.py --list-roles` for the live version of this table directly from the script.

---

## Known Limitations

- **Python 3.10+ required.** The script uses `list[Node]` and `dict[str, int]` type hint syntax introduced in Python 3.10. It will raise a `TypeError` on Python 3.9 and below.

- **Single-page output only.** All nodes and links are rendered on a single draw.io page. Multi-page output (e.g. one page per region or per NE type) is not supported in v1.00.

- **Layout is not deterministic for `spring`.** The `spring` algorithm uses a fixed random seed (`seed=42`) for reproducibility, but results can still vary slightly depending on `networkx` version. User might still need to re-arrange it for better viewing of the topology.

- **No link labels.** Connections are drawn as unlabelled straight lines. Interface names, VLAN IDs, or IP addresses on links are not supported in the current version.

- **No grouped containers.** Nodes are placed as individual icons on a flat canvas. Logical groupings (e.g. drawing a box around all nodes belonging to the same site) are not generated automatically.

- **Column order dependency.** The Nodes sheet is read positionally (column 1 = hostname, 2 = IP, etc.), not by header name. If columns are reordered in the template, the output will be incorrect.

- **Hostname matching is case-sensitive for links.** A link with `source_hostname = JKT-CSR-01` will not match a node named `jkt-csr-01`. Ensure hostname casing is consistent between the Nodes and Links sheets.

- **`networkx` not bundled.** The three graph-based layouts (`kamada_kawai`, `spring`, `spectral`) silently fall back to `hierarchical` if `networkx` is not installed, with a warning message. This may produce unexpected layout results if the user did not intend `hierarchical`.

- - **Layout is static from zero** Program does not support updating topology from existing xml file. Always build layout from ground up.

---

## Changelog

### v1.00 — 2026
- Initial release
- Excel input (`.xlsx`) via `pandas` + `openpyxl`; legacy `.txt` input also supported
- 5 explicit node types: `ME`, `BNG`, `CSR`, `OLT`, `BTS` with pre-assigned Cisco shapes and colors
- Automatic role detection from hostname patterns covering 19 device categories including 5G NFs
- 5 layout algorithms: `kamada_kawai` (default), `spring`, `spectral`, `hierarchical`, `flat`
- Disconnected graph handling — components tiled side-by-side automatically
- Duplicate node and link deduplication with `[warn]` messages
- Unknown link endpoint detection with per-link warning and graceful skip
- Interactive console mode with looping session (generate multiple topologies without restarting)
- CLI mode with full argument flags for scripting and automation
- `--list-roles` utility flag to inspect all node type definitions and patterns
- Custom sheet name support via `--sheet-nodes` and `--sheet-links`
- HTML node labels — 2-line for routers (hostname / IP), 3-line for BTS (tower_id / IP / site_id)
- Output written to `OutputTopo/` directory, auto-created if absent

---

*Flavio Lemuel © 2026 — [github.com/Flalemuel/DrawioGen](https://github.com/Flalemuel/DrawioGen)*
