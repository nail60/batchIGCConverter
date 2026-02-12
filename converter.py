#!/usr/bin/env python3
"""Batch IGC to KMZ converter for paragliding tracklogs."""

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


def convert_file(igc_path, color_kml):
    """Convert a single IGC file to KMZ. Returns (output_path, error_or_None)."""
    name = os.path.splitext(os.path.basename(igc_path))[0]
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


def main():
    # File picker
    files = pick_files_macos()
    if not files:
        sys.exit(0)

    # Color picker
    hex_color = pick_color_macos()
    if hex_color is None:
        sys.exit(0)
    color_kml = rgb_hex_to_kml(hex_color)

    # Convert
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

    # Summary
    msg = f"Converted {success}/{len(files)} files successfully."
    if errors:
        msg += f"\n\n{len(errors)} error(s):\n" + "\n".join(errors)
    show_alert_macos("Conversion Complete", msg)


if __name__ == "__main__":
    main()
