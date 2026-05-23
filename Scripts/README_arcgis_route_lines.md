# Rebuild MiWay Route Lines for ArcGIS Pro

This resets the route-line layer from the GTFS source files in `Data/google_transit`.

## 1. Run the script

From this project folder, run:

```powershell
python Scripts\create_arcgis_route_lines.py
```

If normal `python` is not available, run it from ArcGIS Pro:

1. Open ArcGIS Pro.
2. Open the Python window or a Python Command Prompt.
3. Change to this project folder.
4. Run:

```python
exec(open(r"C:\Users\Raiyan Rizwan\Desktop\UTM Transit Housing Advocacy\Miway Route Efficiency\Scripts\create_arcgis_route_lines.py").read())
```

When run with ArcGIS Pro's Python, the script also tries to create:

```text
Outputs\ArcGIS\miway_route_lines.gdb
```

## 2. Add the line layer

Start with the simpler overview file:

```text
Outputs\ArcGIS\miway_representative_route_lines.geojson
```

Use the detailed version only if you want every GTFS shape variant:

```text
Outputs\ArcGIS\miway_route_shape_lines.geojson
```

## 3. Suggested symbology

Good first map:

- Layer: `miway_representative_route_lines`
- Color by: `score_band`
- Label with: `route_label`
- Optional width by: `advocacy_priority_score`

Good troubleshooting map:

- Layer: `miway_route_shape_lines`
- Color by: `route_short_name`
- Label with: `shape_id`

## 4. What the files contain

Each line has route metadata plus the main efficiency metrics:

- `route_id`
- `route_short_name`
- `route_long_name`
- `route_label`
- `direction_id`
- `shape_id`
- `line_km`
- `trip_count`
- `final_efficiency_score`
- `score_band`
- `advocacy_priority_rank`
- `advocacy_priority_score`
