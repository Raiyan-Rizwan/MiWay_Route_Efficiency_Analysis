"""
Create ArcGIS-ready MiWay route line layers from GTFS shapes.

Outputs:
  - Outputs/ArcGIS/miway_route_shape_lines.geojson
      One feature per unique route/direction/shape_id.
  - Outputs/ArcGIS/miway_representative_route_lines.geojson
      One longest shape per route/direction for a cleaner overview map.
  - Outputs/ArcGIS/miway_route_lines.gdb, when run with ArcGIS Pro's Python.

Run from the project root:
  python Scripts/create_arcgis_route_lines.py

Or from ArcGIS Pro's Python:
  "C:\\Program Files\\ArcGIS\\Pro\\bin\\Python\\envs\\arcgispro-py3\\python.exe" Scripts\\create_arcgis_route_lines.py
"""

from __future__ import annotations

import argparse
import csv
import json
import math
from collections import Counter, defaultdict
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
GTFS_DIR = PROJECT_ROOT / "Data" / "google_transit"
OUTPUT_DIR = PROJECT_ROOT / "Outputs" / "ArcGIS"
METRICS_PATH = PROJECT_ROOT / "Outputs" / "final_route_efficiency_ranking_2026-05-05.csv"


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def safe_float(value: str | None) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except ValueError:
        return None


def haversine_km(lon1: float, lat1: float, lon2: float, lat2: float) -> float:
    radius_km = 6371.0088
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    delta_phi = math.radians(lat2 - lat1)
    delta_lambda = math.radians(lon2 - lon1)
    a = (
        math.sin(delta_phi / 2) ** 2
        + math.cos(phi1) * math.cos(phi2) * math.sin(delta_lambda / 2) ** 2
    )
    return radius_km * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def line_length_km(coords: list[list[float]]) -> float:
    total = 0.0
    for start, end in zip(coords, coords[1:]):
        total += haversine_km(start[0], start[1], end[0], end[1])
    return total


def load_shapes() -> dict[str, list[list[float]]]:
    points_by_shape: dict[str, list[tuple[int, float, float]]] = defaultdict(list)
    for row in read_csv(GTFS_DIR / "shapes.txt"):
        shape_id = row["shape_id"].strip()
        lat = safe_float(row.get("shape_pt_lat"))
        lon = safe_float(row.get("shape_pt_lon"))
        sequence = int(float(row["shape_pt_sequence"]))
        if lat is None or lon is None:
            continue
        points_by_shape[shape_id].append((sequence, lon, lat))

    lines: dict[str, list[list[float]]] = {}
    for shape_id, points in points_by_shape.items():
        ordered = sorted(points, key=lambda item: item[0])
        coords = [[lon, lat] for _, lon, lat in ordered]
        if len(coords) >= 2:
            lines[shape_id] = coords
    return lines


def load_route_lookup() -> dict[str, dict[str, str]]:
    lookup = {}
    for row in read_csv(GTFS_DIR / "routes.txt"):
        route_id = row["route_id"].strip()
        short_name = row.get("route_short_name", "").strip()
        long_name = row.get("route_long_name", "").strip()
        lookup[route_id] = {
            "route_id": route_id,
            "route_short_name": short_name,
            "route_long_name": long_name,
            "route_label": f"{short_name} {long_name}".strip(),
            "route_color": row.get("route_color", "").strip(),
            "route_text_color": row.get("route_text_color", "").strip(),
        }
    return lookup


def load_metrics(metrics_path: Path) -> dict[str, dict[str, object]]:
    if not metrics_path.exists():
        return {}

    keep_columns = {
        "efficiency_rank",
        "final_efficiency_score",
        "score_band",
        "advocacy_priority_rank",
        "advocacy_priority_score",
        "avg_scheduled_speed_kmh",
        "median_headway_min",
        "avg_directness_ratio",
        "stops_per_km",
        "full_service_span_hours",
    }
    metrics = {}
    for row in read_csv(metrics_path):
        route_id = row.get("route_id", "").strip()
        if not route_id:
            continue
        clean = {}
        for column in keep_columns:
            value = row.get(column, "")
            number = safe_float(value)
            clean[column] = number if number is not None else value
        metrics[route_id] = clean
    return metrics


def shape_route_assignments() -> dict[tuple[str, str, str], int]:
    counter: Counter[tuple[str, str, str]] = Counter()
    for row in read_csv(GTFS_DIR / "trips.txt"):
        route_id = row.get("route_id", "").strip()
        direction_id = row.get("direction_id", "").strip()
        shape_id = row.get("shape_id", "").strip()
        if route_id and shape_id:
            counter[(route_id, direction_id, shape_id)] += 1
    return dict(counter)


def build_shape_features() -> list[dict[str, object]]:
    shapes = load_shapes()
    routes = load_route_lookup()
    metrics = load_metrics(METRICS_PATH)
    assignments = shape_route_assignments()

    features = []
    for (route_id, direction_id, shape_id), trip_count in sorted(assignments.items()):
        coords = shapes.get(shape_id)
        if not coords:
            continue
        properties = {
            "route_id": route_id,
            "direction_id": direction_id,
            "shape_id": shape_id,
            "trip_count": trip_count,
            "point_count": len(coords),
            "line_km": round(line_length_km(coords), 3),
        }
        properties.update(routes.get(route_id, {}))
        properties.update(metrics.get(route_id, {}))
        features.append(
            {
                "type": "Feature",
                "properties": properties,
                "geometry": {"type": "LineString", "coordinates": coords},
            }
        )
    return features


def choose_representative_features(
    shape_features: list[dict[str, object]]
) -> list[dict[str, object]]:
    best_by_route_direction: dict[tuple[str, str], dict[str, object]] = {}
    for feature in shape_features:
        properties = feature["properties"]
        key = (str(properties["route_id"]), str(properties["direction_id"]))
        current = best_by_route_direction.get(key)
        if current is None:
            best_by_route_direction[key] = feature
            continue
        current_properties = current["properties"]
        if (
            float(properties["line_km"]),
            int(properties["trip_count"]),
            int(properties["point_count"]),
        ) > (
            float(current_properties["line_km"]),
            int(current_properties["trip_count"]),
            int(current_properties["point_count"]),
        ):
            best_by_route_direction[key] = feature

    representative = []
    for feature in best_by_route_direction.values():
        properties = dict(feature["properties"])
        properties["representative_line"] = 1
        representative.append(
            {
                "type": "Feature",
                "properties": properties,
                "geometry": feature["geometry"],
            }
        )
    return sorted(
        representative,
        key=lambda item: (
            str(item["properties"].get("route_short_name", "")),
            str(item["properties"].get("direction_id", "")),
        ),
    )


def write_geojson(path: Path, features: list[dict[str, object]]) -> None:
    collection = {
        "type": "FeatureCollection",
        "name": path.stem,
        "crs": {
            "type": "name",
            "properties": {"name": "urn:ogc:def:crs:OGC:1.3:CRS84"},
        },
        "features": features,
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(collection, indent=2), encoding="utf-8")


def add_field(arcpy, feature_class: str, name: str, field_type: str) -> None:
    existing = {field.name for field in arcpy.ListFields(feature_class)}
    if name not in existing:
        arcpy.management.AddField(feature_class, name, field_type)


def create_feature_class_with_arcpy(
    features: list[dict[str, object]], feature_class_name: str
) -> str | None:
    try:
        import arcpy  # type: ignore
    except ImportError:
        return None

    gdb_path = OUTPUT_DIR / "miway_route_lines.gdb"
    if not gdb_path.exists():
        arcpy.management.CreateFileGDB(str(OUTPUT_DIR), gdb_path.name)

    feature_class = str(gdb_path / feature_class_name)
    if arcpy.Exists(feature_class):
        arcpy.management.Delete(feature_class)

    spatial_reference = arcpy.SpatialReference(4326)
    arcpy.management.CreateFeatureclass(
        str(gdb_path),
        feature_class_name,
        "POLYLINE",
        spatial_reference=spatial_reference,
    )

    field_types = {
        "route_id": "TEXT",
        "route_short_name": "TEXT",
        "route_long_name": "TEXT",
        "route_label": "TEXT",
        "direction_id": "TEXT",
        "shape_id": "TEXT",
        "route_color": "TEXT",
        "route_text_color": "TEXT",
        "score_band": "TEXT",
        "trip_count": "LONG",
        "point_count": "LONG",
        "representative_line": "SHORT",
        "line_km": "DOUBLE",
        "efficiency_rank": "DOUBLE",
        "final_efficiency_score": "DOUBLE",
        "advocacy_priority_rank": "DOUBLE",
        "advocacy_priority_score": "DOUBLE",
        "avg_scheduled_speed_kmh": "DOUBLE",
        "median_headway_min": "DOUBLE",
        "avg_directness_ratio": "DOUBLE",
        "stops_per_km": "DOUBLE",
        "full_service_span_hours": "DOUBLE",
    }
    for field_name, field_type in field_types.items():
        add_field(arcpy, feature_class, field_name, field_type)

    insert_fields = ["SHAPE@"] + list(field_types)
    with arcpy.da.InsertCursor(feature_class, insert_fields) as cursor:
        for feature in features:
            coords = feature["geometry"]["coordinates"]
            array = arcpy.Array([arcpy.Point(lon, lat) for lon, lat in coords])
            polyline = arcpy.Polyline(array, spatial_reference)
            properties = feature["properties"]
            row = [polyline] + [properties.get(field) for field in field_types]
            cursor.insertRow(row)

    return feature_class


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create MiWay route line layers from GTFS for ArcGIS Pro."
    )
    parser.add_argument(
        "--skip-gdb",
        action="store_true",
        help="Only write GeoJSON files; do not try to create an ArcGIS geodatabase.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    shape_features = build_shape_features()
    representative_features = choose_representative_features(shape_features)

    shape_geojson = OUTPUT_DIR / "miway_route_shape_lines.geojson"
    representative_geojson = OUTPUT_DIR / "miway_representative_route_lines.geojson"
    write_geojson(shape_geojson, shape_features)
    write_geojson(representative_geojson, representative_features)

    print(f"Wrote {len(shape_features):,} route-shape line features:")
    print(f"  {shape_geojson}")
    print(f"Wrote {len(representative_features):,} representative route-direction features:")
    print(f"  {representative_geojson}")

    if args.skip_gdb:
        return

    shape_fc = create_feature_class_with_arcpy(shape_features, "miway_route_shape_lines")
    representative_fc = create_feature_class_with_arcpy(
        representative_features, "miway_representative_route_lines"
    )
    if shape_fc and representative_fc:
        print("Created ArcGIS Pro feature classes:")
        print(f"  {shape_fc}")
        print(f"  {representative_fc}")
    else:
        print("arcpy was not available, so only GeoJSON files were created.")
        print("Run this with ArcGIS Pro's Python to also create the file geodatabase.")


if __name__ == "__main__":
    main()
