#!/usr/bin/env python3
"""Batch IGC to KMZ converter for paragliding tracklogs."""

import colorsys
import os
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime


def parse_igc(filepath):
    """Parse an IGC file and return metadata + tracklog points."""
    points = []
    pilot = None
    date = None

    with open(filepath, "r", errors="replace") as f:
        for line in f:
            line = line.strip()

            # H-records: metadata
            if line.startswith("H"):
                upper = line.upper()
                # Pilot name
                if "PLT" in upper and ":" in line:
                    pilot = line.split(":", 1)[1].strip() or None
                # Date (HFDTE or HPDTE): DDMMYY
                if upper.startswith("HFDTE") or upper.startswith("HPDTE"):
                    digits = "".join(c for c in line[5:] if c.isdigit())
                    if len(digits) >= 6:
                        try:
                            date = datetime.strptime(digits[:6], "%d%m%y").strftime("%Y-%m-%d")
                        except ValueError:
                            pass

            # B-records: fixes
            # BHHMMSSDDMMmmmNDDDMMmmmEVPPPPPGGGGG
            if line.startswith("B") and len(line) >= 35:
                validity = line[24]
                if validity != "A":
                    continue

                try:
                    # Latitude: DDMMmmm
                    lat_deg = int(line[7:9])
                    lat_min = int(line[9:14]) / 1000.0
                    lat = lat_deg + lat_min / 60.0
                    if line[14] == "S":
                        lat = -lat

                    # Longitude: DDDMMmmm
                    lon_deg = int(line[15:18])
                    lon_min = int(line[18:23]) / 1000.0
                    lon = lon_deg + lon_min / 60.0
                    if line[23] == "W":
                        lon = -lon

                    # Altitude: GPS preferred (cols 30-35), pressure fallback (25-30)
                    gps_alt = int(line[30:35])
                    press_alt = int(line[25:30])
                    alt = gps_alt if gps_alt > 0 else press_alt

                    points.append((lat, lon, alt))
                except (ValueError, IndexError):
                    continue

    return {"pilot": pilot, "date": date, "points": points}


def rgb_hex_to_kml(hex_color, alpha=255):
    """Convert '#RRGGBB' hex to KML's AABBGGRR format."""
    hex_color = hex_color.lstrip("#")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return f"{alpha:02x}{b:02x}{g:02x}{r:02x}"


def build_kml(track_data, name, color_kml):
    """Build KML XML for a single flight track."""
    kml = ET.Element("kml", xmlns="http://www.opengis.net/kml/2.2")
    doc = ET.SubElement(kml, "Document")
    ET.SubElement(doc, "name").text = name

    # Description
    parts = []
    if track_data["pilot"]:
        parts.append(f"Pilot: {track_data['pilot']}")
    if track_data["date"]:
        parts.append(f"Date: {track_data['date']}")
    if parts:
        ET.SubElement(doc, "description").text = "\n".join(parts)

    # Line style
    style = ET.SubElement(doc, "Style", id="trackStyle")
    line_style = ET.SubElement(style, "LineStyle")
    ET.SubElement(line_style, "color").text = color_kml
    ET.SubElement(line_style, "width").text = "3"

    # Pin style
    pin_style = ET.SubElement(doc, "Style", id="pinStyle")
    icon_style = ET.SubElement(pin_style, "IconStyle")
    ET.SubElement(icon_style, "scale").text = "1.0"

    points = track_data["points"]
    coords_str = " ".join(f"{lon},{lat},{alt}" for lat, lon, alt in points)

    # Track placemark
    pm = ET.SubElement(doc, "Placemark")
    ET.SubElement(pm, "name").text = name
    ET.SubElement(pm, "styleUrl").text = "#trackStyle"
    ls = ET.SubElement(pm, "LineString")
    ET.SubElement(ls, "altitudeMode").text = "absolute"
    ET.SubElement(ls, "extrude").text = "0"
    ET.SubElement(ls, "tessellate").text = "1"
    ET.SubElement(ls, "coordinates").text = coords_str

    # Takeoff marker
    if points:
        lat, lon, alt = points[0]
        to = ET.SubElement(doc, "Placemark")
        ET.SubElement(to, "name").text = "Takeoff"
        ET.SubElement(to, "styleUrl").text = "#pinStyle"
        pt = ET.SubElement(to, "Point")
        ET.SubElement(pt, "altitudeMode").text = "absolute"
        ET.SubElement(pt, "coordinates").text = f"{lon},{lat},{alt}"

    # Landing marker
    if len(points) > 1:
        lat, lon, alt = points[-1]
        ld = ET.SubElement(doc, "Placemark")
        ET.SubElement(ld, "name").text = "Landing"
        ET.SubElement(ld, "styleUrl").text = "#pinStyle"
        pt = ET.SubElement(ld, "Point")
        ET.SubElement(pt, "altitudeMode").text = "absolute"
        ET.SubElement(pt, "coordinates").text = f"{lon},{lat},{alt}"

    return ET.ElementTree(kml)


def write_kmz(kml_tree, output_path):
    """Write a KML tree to a KMZ file (zipped KML)."""
    kml_bytes = ET.tostring(kml_tree.getroot(), encoding="unicode", xml_declaration=False)
    kml_bytes = '<?xml version="1.0" encoding="UTF-8"?>\n' + kml_bytes

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("doc.kml", kml_bytes)


def convert_file(igc_path, color_kml, output_path=None):
    """Convert a single IGC file to KMZ. Returns (output_path, error_or_None)."""
    name = os.path.splitext(os.path.basename(igc_path))[0]
    if output_path is None:
        output_path = os.path.splitext(igc_path)[0] + ".kmz"

    try:
        data = parse_igc(igc_path)
        if not data["points"]:
            return output_path, "No valid GPS fixes found"
        kml_tree = build_kml(data, name, color_kml)
        write_kmz(kml_tree, output_path)
        return output_path, None
    except Exception as e:
        return output_path, str(e)


def pick_files_macos():
    """Native macOS file picker via AppleScript."""
    script = (
        'set igcFiles to choose file of type {"igc", "IGC"} '
        'with prompt "Select IGC files to convert" '
        'with multiple selections allowed\n'
        'set posixPaths to {}\n'
        'repeat with f in igcFiles\n'
        '  set end of posixPaths to POSIX path of f\n'
        'end repeat\n'
        'set AppleScript\'s text item delimiters to linefeed\n'
        'return posixPaths as text'
    )
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if result.returncode != 0:
        return []
    return [p.strip() for p in result.stdout.strip().split("\n") if p.strip()]


def pick_color_macos():
    """Native macOS color picker via AppleScript. Returns '#RRGGBB' or None."""
    script = 'choose color default color {65535, 0, 0}'
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if result.returncode != 0:
        return None
    # Output: "65535, 0, 0"
    parts = result.stdout.strip().split(",")
    r = int(parts[0].strip()) // 256
    g = int(parts[1].strip()) // 256
    b = int(parts[2].strip()) // 256
    return f"#{r:02x}{g:02x}{b:02x}"


def show_alert_macos(title, message):
    """Native macOS alert dialog via AppleScript."""
    # Escape backslashes and quotes for AppleScript string
    safe_msg = message.replace("\\", "\\\\").replace('"', '\\"')
    safe_title = title.replace("\\", "\\\\").replace('"', '\\"')
    script = f'display dialog "{safe_msg}" with title "{safe_title}" buttons {{"OK"}} default button "OK"'
    subprocess.run(["osascript", "-e", script], capture_output=True)


def pick_mode_macos():
    """Ask user to select files, a folder, or merge KMZ. Returns 'files', 'folder', or 'merge'."""
    script = (
        'display dialog "Convert individual IGC files, an entire folder, or merge existing KMZ files?" '
        'with title "IGC to KMZ Converter" '
        'buttons {"Select Files", "Select Folder", "Merge KMZ"} '
        'default button "Select Files"'
    )
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if result.returncode != 0:
        return None
    if "Merge KMZ" in result.stdout:
        return "merge"
    if "Select Folder" in result.stdout:
        return "folder"
    return "files"


def pick_folder_macos():
    """Native macOS folder picker via AppleScript. Returns POSIX path or None."""
    script = (
        'set chosenFolder to choose folder with prompt "Select folder containing IGC subfolders"\n'
        'return POSIX path of chosenFolder'
    )
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if result.returncode != 0:
        return None
    return result.stdout.strip()


def generate_colors(n):
    """Return n visually distinct hex colors by stepping around the HSV hue wheel."""
    if n == 0:
        return []
    colors = []
    for i in range(n):
        hue = i / n
        r, g, b = colorsys.hsv_to_rgb(hue, 1.0, 1.0)
        colors.append(f"#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}")
    return colors


def convert_folder(folder_path):
    """Convert all IGC files in subfolders, auto-assigning one color per subfolder."""
    folder_path = folder_path.rstrip("/")
    folder_name = os.path.basename(folder_path)
    parent_dir = os.path.dirname(folder_path)
    output_root = os.path.join(parent_dir, folder_name + "_kmz")

    # Scan for subfolders with IGC files and root-level IGC files
    groups = {}  # subfolder_name -> list of igc paths
    root_files = []
    for entry in sorted(os.listdir(folder_path)):
        entry_path = os.path.join(folder_path, entry)
        if os.path.isdir(entry_path):
            igc_files = sorted(
                os.path.join(entry_path, f)
                for f in os.listdir(entry_path)
                if f.lower().endswith(".igc")
            )
            if igc_files:
                groups[entry] = igc_files
        elif entry.lower().endswith(".igc"):
            root_files.append(entry_path)

    if root_files:
        groups[""] = root_files  # empty string key = root level

    if not groups:
        show_alert_macos("No Files Found", "No IGC files found in the selected folder or its subfolders.")
        return

    # Generate colors — one per group
    group_names = sorted(groups.keys())
    colors = generate_colors(len(group_names))
    color_map = dict(zip(group_names, colors))

    # Create output root
    os.makedirs(output_root, exist_ok=True)

    total = sum(len(files) for files in groups.values())
    success = 0
    errors = []
    count = 0

    for group_name in group_names:
        igc_files = groups[group_name]
        hex_color = color_map[group_name]
        color_kml = rgb_hex_to_kml(hex_color)
        display_name = group_name if group_name else "(root)"
        print(f"\n--- {display_name} [{hex_color}] ---")

        # Create output subfolder
        out_dir = os.path.join(output_root, group_name) if group_name else output_root
        os.makedirs(out_dir, exist_ok=True)

        for igc_path in igc_files:
            count += 1
            basename = os.path.splitext(os.path.basename(igc_path))[0]
            out_path = os.path.join(out_dir, basename + ".kmz")
            print(f"[{count}/{total}] Converting {os.path.basename(igc_path)}...")
            out, err = convert_file(igc_path, color_kml, output_path=out_path)
            if err:
                errors.append(f"{os.path.basename(igc_path)}: {err}")
                print(f"  ERROR: {err}")
            else:
                success += 1
                print(f"  -> {os.path.basename(out)}")

    # Summary
    msg = f"Converted {success}/{total} files successfully.\nOutput: {output_root}"
    if errors:
        msg += f"\n\n{len(errors)} error(s):\n" + "\n".join(errors)
    show_alert_macos("Conversion Complete", msg)


def merge_kmz_folder(folder_path):
    """Merge all KMZ files in subfolders into a single combined KMZ."""
    folder_path = folder_path.rstrip("/")
    folder_name = os.path.basename(folder_path)
    parent_dir = os.path.dirname(folder_path)

    ns = "http://www.opengis.net/kml/2.2"

    # Scan for subfolders with KMZ files and root-level KMZ files
    groups = {}  # subfolder_name -> list of kmz paths
    root_files = []
    for entry in sorted(os.listdir(folder_path)):
        entry_path = os.path.join(folder_path, entry)
        if os.path.isdir(entry_path):
            kmz_files = sorted(
                os.path.join(entry_path, f)
                for f in os.listdir(entry_path)
                if f.lower().endswith(".kmz")
            )
            if kmz_files:
                groups[entry] = kmz_files
        elif entry.lower().endswith(".kmz"):
            root_files.append(entry_path)

    if root_files:
        groups["(root)"] = root_files

    if not groups:
        show_alert_macos("No Files Found", "No KMZ files found in the selected folder or its subfolders.")
        return

    # Build combined KML
    kml = ET.Element("kml", xmlns=ns)
    doc = ET.SubElement(kml, "Document")
    ET.SubElement(doc, "name").text = folder_name

    total_files = 0
    for group_name in sorted(groups.keys()):
        folder_el = ET.SubElement(doc, "Folder")
        ET.SubElement(folder_el, "name").text = group_name

        for kmz_path in groups[group_name]:
            total_files += 1
            file_prefix = os.path.splitext(os.path.basename(kmz_path))[0] + "_"
            print(f"  Adding {os.path.basename(kmz_path)}...")

            try:
                with zipfile.ZipFile(kmz_path, "r") as zf:
                    with zf.open("doc.kml") as kml_file:
                        inner_tree = ET.parse(kml_file)
            except Exception as e:
                print(f"  ERROR reading {os.path.basename(kmz_path)}: {e}")
                continue

            inner_root = inner_tree.getroot()
            # Find Document element (handle namespace)
            inner_doc = inner_root.find(f"{{{ns}}}Document")
            if inner_doc is None:
                inner_doc = inner_root.find("Document")
            if inner_doc is None:
                continue

            # Copy Style elements with prefixed IDs
            for style in inner_doc.findall(f"{{{ns}}}Style"):
                old_id = style.get("id", "")
                style.set("id", file_prefix + old_id)
                folder_el.append(style)

            # Copy Placemarks with updated styleUrl references
            for pm in inner_doc.findall(f"{{{ns}}}Placemark"):
                style_url = pm.find(f"{{{ns}}}styleUrl")
                if style_url is not None and style_url.text and style_url.text.startswith("#"):
                    style_url.text = "#" + file_prefix + style_url.text[1:]
                folder_el.append(pm)

    output_path = os.path.join(parent_dir, folder_name + "_merged.kmz")
    kml_tree = ET.ElementTree(kml)
    write_kmz(kml_tree, output_path)

    msg = f"Merged {total_files} KMZ files into:\n{output_path}"
    print(f"\n{msg}")
    show_alert_macos("Merge Complete", msg)


def main():
    # Mode selection
    mode = pick_mode_macos()
    if mode is None:
        sys.exit(0)

    if mode == "merge":
        folder = pick_folder_macos()
        if not folder:
            sys.exit(0)
        merge_kmz_folder(folder)
    elif mode == "folder":
        folder = pick_folder_macos()
        if not folder:
            sys.exit(0)
        convert_folder(folder)
    else:
        # Existing file flow
        files = pick_files_macos()
        if not files:
            sys.exit(0)

        hex_color = pick_color_macos()
        if hex_color is None:
            sys.exit(0)
        color_kml = rgb_hex_to_kml(hex_color)

        success = 0
        errors = []
        for i, f in enumerate(files, 1):
            print(f"[{i}/{len(files)}] Converting {os.path.basename(f)}...")
            out, err = convert_file(f, color_kml)
            if err:
                errors.append(f"{os.path.basename(f)}: {err}")
                print(f"  ERROR: {err}")
            else:
                success += 1
                print(f"  -> {os.path.basename(out)}")

        msg = f"Converted {success}/{len(files)} files successfully."
        if errors:
            msg += f"\n\n{len(errors)} error(s):\n" + "\n".join(errors)
        show_alert_macos("Conversion Complete", msg)


if __name__ == "__main__":
    main()
