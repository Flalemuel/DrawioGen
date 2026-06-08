#!/usr/bin/env python3
"""
draw.io Network Topology XML Generator
Reads nodes and links from an Excel file and generates a draw.io diagram.

Node types supported (via node_type column):
    ME   — Metro Router       red fill, Cisco router icon
    BNG  — Broadband Gateway  dark blue fill, Cisco router icon
    CSR  — Core/Service Router Cisco blue fill, Cisco router icon
    OLT  — Optical Line Term  pink fill, rectangle icon
    BTS  — Base Tower Station orange fill, Cisco radio tower icon
               label = tower_id / ip / site_id  (3 lines)
    (auto) — fallback role detected from hostname patterns

Excel sheets:
    Nodes  columns: hostname/tower_id | ip_loopback | node_type | site_id
    Links  columns: source_hostname   | destination_hostname

Usage:
    # Default — best for most topologies
        python drawio_topology.py -i topology.xlsx -o out.drawio

    # Explicit layout selection
        python drawio_topology.py -i topology.xlsx -o out.drawio --layout kamada_kawai
        python drawio_topology.py -i topology.xlsx -o out.drawio --layout spring
        python drawio_topology.py -i topology.xlsx -o out.drawio --layout spectral
        python drawio_topology.py -i topology.xlsx -o out.drawio --layout hierarchical
        python drawio_topology.py -i topology.xlsx -o out.drawio --layout flat

        Algorithm	    Best for
        kamada_kawai	Default — balanced, minimises edge length variance, reads like a real network diagram
        spring	        Dense or mesh networks, organic clustering
        spectral	    Very large graphs (100+ nodes), fast
        hierarchical	Strict top-down trees, no networkx needed
        flat	        No links, grouped by device type
"""

import argparse
import os
import sys
import re
import xml.etree.ElementTree as ET
from xml.dom import minidom
from dataclasses import dataclass, field


# ---------------------------------------------------------------------------
# Explicit node-type definitions (node_type column takes priority)
# ---------------------------------------------------------------------------
NODE_TYPES = {
    "ME": {
        "fill":   "#E51400",   # red
        "stroke": "#B20000",
        "font_color": "#FFFFFF",
        "shape":  "shape=mxgraph.cisco.routers.router;",
        "label":  "Metro Router",
        "is_bts": False,
        "w": 78, "h": 53,
    },
    "BNG": {
        "fill":   "#5552FF",   # dark blue
        "stroke": "#036897",
        "font_color": "#FFFFFF",
        "shape":  "shape=mxgraph.cisco.routers.router;",
        "label":  "BNG/BRAS",
        "is_bts": False,
        "w": 78, "h": 60
    },
    "CSR": {
        "fill":   "#036897",   # Cisco blue
        "stroke": "#FFFFFF",
        "font_color": "#FFFFFF",
        "shape":  "shape=mxgraph.cisco.routers.router;",
        "label":  "CSR",
        "is_bts": False,
        "w": 78, "h": 53,
    },
    "OLT": {
        "fill":   "#F8CECC",   # light pink
        "stroke": "#B85450",
        "font_color": "#000000",
        "shape":  "shape=mxgraph.cisco.switches.workgroup_switch;",
        "label":  "OLT/ONU",
        "is_bts": False,
        "w": 60, "h": 40,
    },
    "BTS": {
        "fill":   "#036897",   # light orange
        "stroke": "#036897",
        "font_color": "#000000",
        "shape":  "shape=mxgraph.cisco.wireless.radio_tower;",
        "label":  "BTS",
        "is_bts": True,
        "w": 37, "h": 101,   # taller to fit radio tower icon
    },
}


# ---------------------------------------------------------------------------
# Fallback roles detected from hostname patterns (when node_type is blank) - Differentiate different vendor
# ---------------------------------------------------------------------------
ROLES = [
    { "key":"core",   "patterns":["core","backbone","bb-"],                "fill":"#dae8fc","stroke":"#6c8ebf","font_color":"#000000","shape":"shape=mxgraph.cisco.routers.router;",                   "label":"Core Router",   "w":60,"h":60 },
    { "key":"pe",     "patterns":["pe-","pe_","asbr","edge-rtr"],          "fill":"#d5e8d4","stroke":"#82b366","font_color":"#000000","shape":"shape=mxgraph.cisco.routers.router;",                   "label":"PE Router",     "w":60,"h":60 },
    { "key":"p",      "patterns":["-p-","_p_","transit-"],                 "fill":"#dae8fc","stroke":"#6c8ebf","font_color":"#000000","shape":"shape=mxgraph.cisco.routers.router;",                   "label":"P Router",      "w":60,"h":60 },
    { "key":"upf",    "patterns":["upf"],                                  "fill":"#e1d5e7","stroke":"#9673a6","font_color":"#000000","shape":"shape=mxgraph.cisco.servers.standard_server;",          "label":"UPF",           "w":60,"h":60 },
    { "key":"amf",    "patterns":["amf"],                                  "fill":"#e1d5e7","stroke":"#9673a6","font_color":"#000000","shape":"shape=mxgraph.cisco.servers.standard_server;",          "label":"AMF",           "w":60,"h":60 },
    { "key":"smf",    "patterns":["smf"],                                  "fill":"#e1d5e7","stroke":"#9673a6","font_color":"#000000","shape":"shape=mxgraph.cisco.servers.standard_server;",          "label":"SMF",           "w":60,"h":60 },
    { "key":"pcf",    "patterns":["pcf"],                                  "fill":"#e1d5e7","stroke":"#9673a6","font_color":"#000000","shape":"shape=mxgraph.cisco.servers.standard_server;",          "label":"PCF",           "w":60,"h":60 },
    { "key":"5gc_nf", "patterns":["udm","udr","ausf","nrf","nssf","nef"],  "fill":"#e1d5e7","stroke":"#9673a6","font_color":"#000000","shape":"shape=mxgraph.cisco.servers.standard_server;",          "label":"5GC NF",        "w":60,"h":60 },
    { "key":"me",     "patterns":["-me-","_me_","me1","me2","me3"],        "fill":"#FFA513","stroke":"#AE0000","font_color":"#FFFFFF","shape":"shape=mxgraph.cisco.routers.router;",                   "label":"Metro Router",  "w":60,"h":60 },
    { "key":"gnb",    "patterns":["gnb","enb","ran-","radio-","site-"],    "fill":"#fff2cc","stroke":"#d6b656","font_color":"#000000","shape":"shape=mxgraph.cisco.wireless.wireless_access_point;",   "label":"gNB",           "w":60,"h":60 },
    { "key":"bts",    "patterns":["bts","twr-","tower-"],                  "fill":"#FFE6CC","stroke":"#D6772E","font_color":"#000000","shape":"shape=mxgraph.cisco.wireless.radio_tower;",             "label":"BTS",           "w":60,"h":80 },
    { "key":"olt",    "patterns":["olt","onu","ont","-fh","_fh"],          "fill":"#F8CECC","stroke":"#B85450","font_color":"#000000","shape":"rounded=0;whiteSpace=wrap;",                            "label":"OLT/ONU",       "w":90,"h":50 },
    { "key":"bng",    "patterns":["bng","bras"],                           "fill":"#003087","stroke":"#001A4D","font_color":"#FFFFFF","shape":"shape=mxgraph.cisco.routers.router;",                   "label":"BNG/BRAS",      "w":60,"h":60 },
    { "key":"csr",    "patterns":["csr","asr"],                            "fill":"#1BA0D7","stroke":"#007FAD","font_color":"#FFFFFF","shape":"shape=mxgraph.cisco.routers.router;",                   "label":"CSR",           "w":60,"h":60 },
    { "key":"sw",     "patterns":["sw-","sw_","-sw","switch","agg-","nx"], "fill":"#dae8fc","stroke":"#6c8ebf","font_color":"#000000","shape":"shape=mxgraph.cisco.switches.workgroup_switch;",        "label":"Switch",        "w":60,"h":60 },
    { "key":"fw",     "patterns":["fw","firewall","asa-","palo-"],         "fill":"#f8cecc","stroke":"#b85450","font_color":"#000000","shape":"shape=mxgraph.cisco.firewalls.firewall;",               "label":"Firewall",      "w":60,"h":60 },
    { "key":"cpe",    "patterns":["cpe","home-"],                          "fill":"#fff2cc","stroke":"#d6b656","font_color":"#000000","shape":"shape=mxgraph.cisco.routers.router;",                   "label":"CPE",           "w":60,"h":60 },
    { "key":"server", "patterns":["srv","server","vm-","dc-"],             "fill":"#e1d5e7","stroke":"#9673a6","font_color":"#000000","shape":"shape=mxgraph.cisco.servers.standard_server;",          "label":"Server",        "w":60,"h":60 },
]
DEFAULT_ROLE = { "key":"other","fill":"#f5f5f5","stroke":"#999","font_color":"#000000","shape":"shape=mxgraph.cisco.routers.router;","label":"Device","w":60,"h":60 }
ROLE_ORDER   = ["core","p","pe","bng","csr","me","upf","amf","smf","pcf","5gc_nf","gnb","bts","olt","sw","fw","server","cpe","other"]


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------
@dataclass
class Node:
    hostname: str          # hostname for routers; tower_id for BTS
    ip: str
    node_type: str         # ME/BNG/CSR/OLT/BTS or "" for auto-detect
    site_id: str            # BTS only
    role: dict             # resolved style dict
    x: float = 0.0
    y: float = 0.0
    cell_id: int = 0

@dataclass
class Link:
    source: str
    target: str


# ---------------------------------------------------------------------------
# Role resolution
# ---------------------------------------------------------------------------
def resolve_role(hostname: str, node_type: str) -> dict:
    """Explicit node_type column takes priority over pattern matching."""
    nt = node_type.strip().upper() if node_type else ""
    if nt in NODE_TYPES:
        r = dict(NODE_TYPES[nt])
        r["key"] = nt.lower()
        return r
    # fallback — pattern match on hostname
    h = hostname.lower()
    for role in ROLES:
        if any(p in h for p in role["patterns"]):
            return role
    return DEFAULT_ROLE


def make_label(node: Node) -> str:
    """
    All labels use HTML so line breaks render correctly in draw.io.
    BTS  -> tower_id (bold) / ip / site_id  (3 lines)
    Others -> hostname / ip                 (2 lines)
    """
    nt = node.node_type.strip().upper() if node.node_type else ""
    is_bts = nt == "BTS" or node.role.get("key") in ("bts",)

    def esc(s):
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    if is_bts:
        inner = f"<b>{esc(node.hostname)}</b>"
        if node.ip:     inner += f"<br>{esc(node.ip)}"
        if node.site_id: inner += f"<br><font color='#666666'>{esc(node.site_id)}</font>"
        return f"<html><center>{inner}</center></html>"
    else:
        inner = esc(node.hostname)
        if node.ip: inner += f"<br>{esc(node.ip)}"
        return f"<html><center>{inner}</center></html>"


# ---------------------------------------------------------------------------
# Excel reader
# ---------------------------------------------------------------------------
def read_excel(path: str, sheet_nodes: str, sheet_links: str):
    try:
        import pandas as pd
    except ImportError:
        print("Error: pip install pandas openpyxl", file=sys.stderr); sys.exit(1)
    try:
        xl = pd.read_excel(path, sheet_name=None, dtype=str)
    except FileNotFoundError:
        print(f"Error: file not found: {path}", file=sys.stderr); sys.exit(1)
    except Exception as e:
        print(f"Error reading Excel: {e}", file=sys.stderr); sys.exit(1)

    if sheet_nodes not in xl:
        print(f"Error: sheet '{sheet_nodes}' not found. Got: {list(xl.keys())}", file=sys.stderr); sys.exit(1)

    df = xl[sheet_nodes].dropna(how="all")
    df.columns = [str(c).strip().lower() for c in df.columns]

    nodes, seen = [], set()
    for _, row in df.iterrows():
        hostname  = str(row.iloc[0]).strip() if len(row) > 0 else ""
        ip        = str(row.iloc[1]).strip() if len(row) > 1 else ""
        node_type = str(row.iloc[2]).strip() if len(row) > 2 else ""
        site_id    = str(row.iloc[3]).strip() if len(row) > 3 else ""

        # clean NaN strings
        hostname  = "" if hostname  in ("nan","none","hostname / tower_id","hostname") else hostname
        ip        = "" if ip        in ("nan","none","ip_loopback")  else ip
        node_type = "" if node_type in ("nan","none","node_type")    else node_type
        site_id    = "" if site_id    in ("nan","none","site_id")       else site_id

        if not hostname: continue
        if hostname in seen:
            print(f"  [warn] duplicate skipped: {hostname}", file=sys.stderr); continue
        seen.add(hostname)
        nodes.append(Node(
            hostname=hostname, ip=ip,
            node_type=node_type, site_id=site_id,
            role=resolve_role(hostname, node_type),
        ))

    links, seen_pairs = [], set()
    if sheet_links in xl:
        df2 = xl[sheet_links].dropna(how="all")
        for _, row in df2.iterrows():
            src = str(row.iloc[0]).strip() if len(row) > 0 else ""
            dst = str(row.iloc[1]).strip() if len(row) > 1 else ""
            src = "" if src in ("nan","none","source_hostname") else src
            dst = "" if dst in ("nan","none","destination_hostname") else dst
            if not src or not dst: continue
            if (src, dst) not in seen_pairs:
                seen_pairs.add((src, dst))
                links.append(Link(source=src, target=dst))
    else:
        print(f"  [info] No Links sheet '{sheet_links}' — nodes-only diagram.", file=sys.stderr)

    return nodes, links


# ---------------------------------------------------------------------------
# Text file readers (legacy)
# ---------------------------------------------------------------------------
def parse_nodes_text(text: str) -> list[Node]:
    nodes, seen = [], set()
    for lineno, line in enumerate(text.splitlines(), 1):
        line = line.strip()
        if not line or line.startswith("#"): continue
        parts = re.split(r"[\s,;|]+", line)
        if len(parts) < 2:
            print(f"  [warn] line {lineno} skipped", file=sys.stderr); continue
        hostname, ip = parts[0], parts[1]
        node_type = parts[2] if len(parts) > 2 else ""
        site_id    = parts[3] if len(parts) > 3 else ""
        if hostname in seen: continue
        seen.add(hostname)
        nodes.append(Node(hostname=hostname, ip=ip, node_type=node_type, site_id=site_id,
                          role=resolve_role(hostname, node_type)))
    return nodes

def parse_links_text(text: str) -> list[Link]:
    links, seen = [], set()
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"): continue
        parts = re.split(r"[\s,;|]+", line)
        if len(parts) < 2: continue
        if (parts[0], parts[1]) not in seen:
            seen.add((parts[0], parts[1]))
            links.append(Link(parts[0], parts[1]))
    return links


# ---------------------------------------------------------------------------
# Layout — hierarchical or flat grid
# ---------------------------------------------------------------------------
def layout_nodes(nodes: list[Node], links: list[Link], algo: str = "kamada_kawai") -> None:
    """
    Layout algorithms (--layout flag):
      kamada_kawai  — balanced, minimises edge length variance (default)
      spring        — force-directed, organic feel, good for dense graphs
      spectral      — eigenvector-based, fast for large graphs
      hierarchical  — original top-down tree layout (no networkx needed)
      flat          — flat grid grouped by role (no links needed)
    """
    if algo == "flat" or not links:
        _layout_flat(nodes); return
    if algo == "hierarchical":
        _layout_hierarchical(nodes, links); return

    try:
        import networkx as nx
    except ImportError:
        print("  [warn] networkx not found — falling back to hierarchical layout.", file=sys.stderr)
        print("         Install with: pip install networkx", file=sys.stderr)
        _layout_hierarchical(nodes, links); return

    # Canvas size — scale factor spreads nodes out proportionally
    SCALE     = 220   # pixels between nodes (increase to spread further)
    MARGIN    = 100   # canvas margin
    NODE_W    = 60    # reference node width for centre-offset

    hset = {n.hostname for n in nodes}

    # Build networkx graph from valid links only
    G = nx.Graph()
    G.add_nodes_from(hset)
    for lk in links:
        if lk.source in hset and lk.target in hset:
            G.add_edge(lk.source, lk.target)

    # Compute positions — networkx returns normalised coords in [-1, 1]
    if algo == "kamada_kawai":
        # Best for network topology — minimises edge length differences
        # Handle disconnected graphs by computing per-component then merging
        if nx.is_connected(G):
            pos = nx.kamada_kawai_layout(G, scale=1.0)
        else:
            pos = _layout_disconnected(G, nx.kamada_kawai_layout, scale=1.0)

    elif algo == "spring":
        # Force-directed — good for organic/mesh topologies
        # k controls ideal edge length; higher = more spread
        pos = nx.spring_layout(G, k=2.5 / (len(nodes) ** 0.5), iterations=100, seed=42)

    elif algo == "spectral":
        # Eigenvector-based — fast for large graphs
        if nx.is_connected(G):
            pos = nx.spectral_layout(G, scale=1.0)
        else:
            pos = _layout_disconnected(G, nx.spectral_layout, scale=1.0)
    else:
        print(f"  [warn] unknown layout '{algo}', using kamada_kawai", file=sys.stderr)
        pos = nx.kamada_kawai_layout(G, scale=1.0)

    # Map normalised [-1,1] coords to canvas pixel coordinates
    # Find bounding box of computed positions
    xs = [p[0] for p in pos.values()]
    ys = [p[1] for p in pos.values()]
    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)
    span_x = max_x - min_x or 1.0
    span_y = max_y - min_y or 1.0

    node_map = {n.hostname: n for n in nodes}
    for hostname, (nx_x, nx_y) in pos.items():
        node = node_map.get(hostname)
        if not node: continue
        w = node.role.get("w", NODE_W)
        h = node.role.get("h", 60)
        # Normalise to [0,1] then scale to canvas, offset by half node size
        node.x = MARGIN + ((nx_x - min_x) / span_x) * SCALE * (len(nodes) ** 0.5) - w / 2
        node.y = MARGIN + ((nx_y - min_y) / span_y) * SCALE * (len(nodes) ** 0.5) - h / 2

    # Any nodes not in the graph (isolated, not in links) — append below canvas
    placed = set(pos.keys())
    orphans = [n for n in nodes if n.hostname not in placed]
    if orphans:
        max_y_canvas = max(n.y for n in nodes if n.hostname in placed) + 160
        for i, node in enumerate(orphans):
            node.x = MARGIN + i * 180
            node.y = max_y_canvas


def _layout_disconnected(G, layout_fn, **kwargs):
    """Apply layout per connected component then tile components side by side."""
    import networkx as nx
    pos = {}
    x_offset = 0.0
    for component in nx.connected_components(G):
        subgraph = G.subgraph(component)
        if len(subgraph) == 1:
            node = list(subgraph.nodes)[0]
            sub_pos = {node: (x_offset, 0.0)}
        else:
            sub_pos = layout_fn(subgraph, **kwargs)
        # shift component to sit next to previous one
        if sub_pos:
            min_x = min(p[0] for p in sub_pos.values())
            shifted = {n: (x_offset + (p[0] - min_x), p[1]) for n, p in sub_pos.items()}
            max_x = max(p[0] for p in shifted.values())
            x_offset = max_x + 0.4   # gap between components
            pos.update(shifted)
    return pos


def _layout_hierarchical(nodes: list[Node], links: list[Link]) -> None:
    """Original top-down hierarchical tree layout (no external libraries)."""
    CG, RG, SX, SY = 120, 140, 60, 60
    hset     = {n.hostname for n in nodes}
    parents  = {n.hostname: set() for n in nodes}
    children = {n.hostname: set() for n in nodes}
    for lk in links:
        if lk.source in hset and lk.target in hset:
            children[lk.source].add(lk.target)
            parents[lk.target].add(lk.source)
    level: dict[str, int] = {}
    def assign(h, visited):
        if h in level: return level[h]
        if h in visited: level[h] = 0; return 0
        visited.add(h)
        level[h] = 0 if not parents[h] else max(assign(p, visited) for p in parents[h]) + 1
        return level[h]
    for n in nodes: assign(n.hostname, set())
    rows: dict[int, list[Node]] = {}
    for n in nodes:
        rows.setdefault(level.get(n.hostname, 0), []).append(n)
    max_cols = max(len(r) for r in rows.values())
    canvas_w = max_cols * (80 + CG)
    for lvl, rn in sorted(rows.items()):
        rw = sum(n.role.get("w", 60) for n in rn) + (len(rn) - 1) * CG
        xs = SX + (canvas_w - rw) / 2
        x  = xs
        for node in rn:
            node.x = x
            node.y = SY + lvl * (80 + RG)
            x += node.role.get("w", 60) + CG


def _layout_flat(nodes: list[Node], cols: int = 6) -> None:
    """Flat grid grouped by role — used when no links provided."""
    CG, RG, SX, SY, SG = 120, 100, 60, 60, 60
    groups: dict[str, list] = {}
    for n in nodes:
        groups.setdefault(n.role.get("key", "other"), []).append(n)
    keys = [k for k in ROLE_ORDER if k in groups] + [k for k in groups if k not in ROLE_ORDER]
    y = SY
    for k in keys:
        rn = groups[k]
        x  = SX
        row_h = 0
        for i, node in enumerate(rn):
            if i > 0 and i % cols == 0:
                y += row_h + RG; x = SX; row_h = 0
            node.x = x; node.y = y
            x += node.role.get("w", 60) + CG
            row_h = max(row_h, node.role.get("h", 60))
        y += row_h + SG


# ---------------------------------------------------------------------------
# XML builder
# ---------------------------------------------------------------------------
def build_xml(nodes: list[Node], links: list[Link], algo: str = "kamada_kawai") -> str:
    layout_nodes(nodes, links, algo)

    for i, node in enumerate(nodes):
        node.cell_id = i + 10

    node_map = {n.hostname: n for n in nodes}

    root = ET.Element("mxGraphModel", {
        "dx":"1422","dy":"762","grid":"1","gridSize":"10",
        "guides":"1","tooltips":"1","connect":"1","arrows":"1",
        "fold":"1","page":"1","pageScale":"1",
        "pageWidth":"1654","pageHeight":"1169","math":"0","shadow":"0",
    })
    root_el = ET.SubElement(root, "root")
    ET.SubElement(root_el, "mxCell", {"id":"0"})
    ET.SubElement(root_el, "mxCell", {"id":"1","parent":"0"})

    for node in nodes:
        r   = node.role
        w   = r.get("w", 60)
        h   = r.get("h", 60)
        nt  = node.node_type.strip().upper() if node.node_type else ""
        is_bts = nt == "BTS" or r.get("key") in ("bts",)
        is_olt = nt == "OLT" or r.get("key") in ("olt",)

        label = make_label(node)

        # html=1 on ALL nodes so HTML labels render correctly.
        # fontColor always #000000 — Cisco shape labels render below the
        # icon as a separate text block; white fontColor in the style is
        # ignored by draw.io causing invisible text on light backgrounds.
        if is_olt:
            style = (
                f"shape=mxgraph.cisco.optical.optical_transport;"
                f"fillColor={r['fill']};strokeColor={r['stroke']};"
                f"fontColor=#000000;fontSize=10;fontStyle=0;"
                f"verticalLabelPosition=bottom;verticalAlign=top;"
                f"html=1;"
            )
        elif is_bts:
            style = (
                f"{r['shape']}"
                f"fillColor={r['fill']};strokeColor={r['stroke']};"
                f"fontColor=#000000;fontSize=9;fontStyle=0;"
                f"verticalLabelPosition=bottom;verticalAlign=top;"
                f"html=1;"
            )
        else:
            style = (
                f"{r['shape']}"
                f"fillColor={r['fill']};strokeColor={r['stroke']};"
                f"fontColor=#000000;fontSize=10;fontStyle=0;"
                f"verticalLabelPosition=bottom;verticalAlign=top;"
                f"html=1;"
            )

        cell = ET.SubElement(root_el, "mxCell", {
            "id":     str(node.cell_id),
            "value":  label,
            "style":  style,
            "vertex": "1",
            "parent": "1",
        })
        ET.SubElement(cell, "mxGeometry", {
            "x": str(int(node.x)), "y": str(int(node.y)),
            "width": str(w), "height": str(h),
            "as": "geometry",
        })

    # Straight lines, perimeter-snapping (no exitX/Y or entryX/Y)
    edge_style = "endArrow=none;edgeStyle=none;rounded=0;strokeColor=#000000;strokeWidth=1;"
    skipped = 0
    edge_id = max(n.cell_id for n in nodes) + 1
    for lk in links:
        src = node_map.get(lk.source)
        dst = node_map.get(lk.target)
        if not src:
            print(f"  [warn] link skipped — unknown source: {lk.source}", file=sys.stderr)
            skipped += 1; continue
        if not dst:
            print(f"  [warn] link skipped — unknown dest: {lk.target}", file=sys.stderr)
            skipped += 1; continue
        cell = ET.SubElement(root_el, "mxCell", {
            "id": str(edge_id), "value": "",
            "style": edge_style, "edge": "1",
            "source": str(src.cell_id), "target": str(dst.cell_id), "parent": "1",
        })
        ET.SubElement(cell, "mxGeometry", {"relative":"1","as":"geometry"})
        edge_id += 1

    if skipped:
        print(f"  [warn] {skipped} link(s) skipped", file=sys.stderr)

    raw = ET.tostring(root, encoding="unicode")
    return minidom.parseString(raw).toprettyxml(indent="  ", encoding=None)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Generate draw.io topology XML from Excel.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python drawio_topology.py -i topology_input.xlsx -o topology.drawio
  python drawio_topology.py --list-roles
  python drawio_topology.py -n nodes.txt -l links.txt -o topology.drawio

node_type column values:
  ME   red router          BNG  dark-blue router
  CSR  Cisco-blue router   OLT  pink rectangle
  BTS  radio tower icon, label = tower_id / ip / site_id
  (blank) = auto-detect from hostname patterns
        """
    )
    parser.add_argument("-i","--input",         help="Excel input (.xlsx)",         default=None)
    parser.add_argument("-o","--output",        help="Output .drawio",              default="topology.drawio")
    parser.add_argument("--sheet-nodes",        help="Nodes sheet name",            default="Nodes")
    parser.add_argument("--sheet-links",        help="Links sheet name",            default="Links")
    parser.add_argument("-n","--nodes",         help="[legacy] nodes text file",    default=None)
    parser.add_argument("-l","--links",         help="[legacy] links text file",    default=None)
    parser.add_argument("--layout", default="kamada_kawai",
        choices=["kamada_kawai","spring","spectral","hierarchical","flat"],
        help="Layout algorithm (default: kamada_kawai)")
    parser.add_argument("--list-roles",         action="store_true")
    args = parser.parse_args()

    if args.list_roles:
        print(f"\nExplicit node_type values (column 3 in Nodes sheet):")
        print(f"  {'Type':<6}  {'Label':<16}  Color/Icon")
        print(f"  {'-'*50}")
        for k,v in NODE_TYPES.items():
            print(f"  {k:<6}  {v['label']:<16}  fill={v['fill']}  {v['shape'][:40]}")
        print(f"\nFallback patterns (when node_type blank):")
        print(f"  {'Key':<10} {'Label':<16} Patterns")
        print(f"  {'-'*60}")
        for r in ROLES:
            print(f"  {r['key']:<10} {r['label']:<16} {', '.join(r['patterns'])}")
        return

    if args.input:
        nodes, links = read_excel(args.input, args.sheet_nodes, args.sheet_links)
    elif args.nodes:
        with open(args.nodes,"r",encoding="utf-8") as f: nodes = parse_nodes_text(f.read())
        links = []
        if args.links:
            with open(args.links,"r",encoding="utf-8") as f: links = parse_links_text(f.read())
    else:
        print("Error: provide -i <excel> or -n <nodes.txt>", file=sys.stderr)
        parser.print_help(); sys.exit(1)

    if not nodes:
        print("Error: no valid nodes found.", file=sys.stderr); sys.exit(1)

    xml_output = build_xml(nodes, links, algo=args.layout)
    with open(args.output,"w",encoding="utf-8") as f: f.write(xml_output)

    role_counts: dict[str,int] = {}
    for n in nodes:
        role_counts[n.role["label"]] = role_counts.get(n.role["label"],0) + 1
    valid = sum(1 for lk in links
                if any(n.hostname==lk.source for n in nodes)
                and any(n.hostname==lk.target for n in nodes))

    print(f"Generated  : {args.output}")
    print(f"Nodes      : {len(nodes)}")
    print(f"Links      : {valid}" + (f" ({len(links)-valid} skipped)" if valid<len(links) else ""))
    print(f"Layout     : {args.layout}")
    print("Breakdown  :")
    for label,count in sorted(role_counts.items()):
        print(f"  {label:<22} {count}")

LAYOUTS = [
    ("kamada_kawai",  "Default — balanced, best for most topologies"),
    ("spring",        "Dense or mesh networks, organic clustering   <-- Best layout"),
    ("spectral",      "Very large graphs (100 nodes), fast"),
    ("hierarchical",  "Strict top-down trees, no networkx needed"),
    ("flat",          "No links, grouped by device type"),
]

BANNER = r"""
  ____                       _           _____
 |  _ \ _ __ __ ___         (_)   ___   |_   _|__  _ __   ___
 | | | | '__/ _` \ \ /\ / / | |  / _ \    | |/ _ \| '_ \ / _ \
 | |_| | | | (_| |\ V  V /  | | | (_) |   | | (_) | |_) | (_) |
 |____/|_|  \__,_| \_/\_/   |_|  \___/    |_|\___/| .__/ \___/
                                                  |_|
        draw.io Network Topology XML Generator
=========================================================
"""

def get_exe_dir():
    """Return the folder where the EXE (or .py script) is located."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def prompt_input_file(base_dir):
    """Prompt until a valid .xlsx file is found in the input subfolder."""
    input_dir = os.path.join(base_dir, "TopoExcel")

    while True:
        print()
        raw = input("  Enter Excel input filename (e.g. topology.xlsx): ").strip()
        if not raw:
            print("  [!] Filename cannot be empty. Please try again.")
            continue

        if not raw.lower().endswith(".xlsx"):
            raw += ".xlsx"

        full_path = os.path.join(input_dir, raw)

        if not os.path.isfile(full_path):
            print(f"  [!] File not found: {full_path}")
            print("      Make sure the file is inside the 'TopoExcel' folder.")
            continue

        return full_path

def prompt_output_name():
    """Prompt for output filename, auto-append .drawio."""
    while True:
        print()
        raw = input("  Enter output filename (without extension, e.g. my_topology): ").strip()
        if not raw:
            print("  [!] Output name cannot be empty. Please try again.")
            continue
        if raw.lower().endswith(".drawio"):
            raw = raw[:-7]
        return raw + ".drawio"

def prompt_layout():
    """Show numbered layout menu and return chosen layout string."""
    print()
    print("  Select layout algorithm:")
    print()
    for i, (name, desc) in enumerate(LAYOUTS, 1):
        default_tag = "  <-- default" if i == 1 else ""
        print(f"    {i}. {name:<16} {desc}{default_tag}")
    print()
    while True:
        choice = input("  Enter choice [1-5] or press Enter for default (1): ").strip()
        if choice == "":
            return LAYOUTS[0][0]
        if choice.isdigit() and 1 <= int(choice) <= len(LAYOUTS):
            return LAYOUTS[int(choice) - 1][0]
        print(f"  [!] Invalid choice. Please enter a number between 1 and {len(LAYOUTS)}.")

def run_interactive():
    """Interactive console session — calls original main() via sys.argv injection."""
    base_dir = get_exe_dir()

    while True:
        print(BANNER)
        print(f"  Working folder: {base_dir}")
        print()

        input_file  = prompt_input_file(base_dir)
        output_name = prompt_output_name()
        layout      = prompt_layout()
        output_dir = os.path.join(base_dir, "OutputTopo")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, output_name)

        print()
        print("  -------------------------------------------------------")
        print(f"  Input  : {input_file}")
        print(f"  Output : {output_path}")
        print(f"  Layout : {layout}")
        print("  -------------------------------------------------------")
        print()
        print("  Generating topology diagram ...")
        print()

        sys.argv = [
            sys.argv[0],
            "-i", input_file,
            "-o", output_path,
            "--layout", layout,
        ]

        try:
            main()
            print()
            print(f"  [OK] Successfully created: {output_path}")
        except SystemExit as e:
            if e.code not in (None, 0):
                print()
                print(f"  [ERROR] The generator exited with an error (code {e.code}).")
                print("          Check the messages above for details.")
        except Exception as e:
            print()
            print(f"  [ERROR] Unexpected error: {e}")

        print()
        again = input("  Generate another topology? [y/n]: ").strip().lower()
        if again not in ("y", "yes"):
            print()
            print("  Goodbye!")
            print()
            break

if __name__ == "__main__":
    if len(sys.argv) > 1:
        main()
    else:
        try:
            run_interactive()
        except KeyboardInterrupt:
            print("\n\n  Interrupted. Exiting.")
        input("\n  Press Enter to close...")