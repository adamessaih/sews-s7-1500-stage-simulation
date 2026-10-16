# SEWS S7-1500 PLC Simulation and Monitoring

Portfolio project developed during an industrial supervision and traceability internship. This repository contains a small, selected set of source files for a **simulation/demo** of a Siemens S7-1500 line, with a Python bridge that exchanges PLC status and commands over MQTT.

## Included

- `TIA_Sources/`: SCL data blocks and simulation function blocks for TIA Portal.
- `Python_Bridge/`: Python Snap7/MQTT bridge, a DB100 connectivity check, dependency list, and an example configuration.

Reports, real production records, operator data, credentials, Node-RED runtime state, compiled TIA/Unity projects, and generated build files are intentionally omitted.

## Architecture

```text
MQTT client / dashboard <-> Python bridge <-> S7-1500 DB100 (TIA Portal / PLCSIM Advanced)
```

The SCL sources model a sequential harness-production demo with manual and automatic modes, status reporting, timing, and fault/alarm states. The bridge reads and writes the DB100 interface and publishes/subscribes to MQTT topics.

## Requirements

- Siemens TIA Portal with S7-1500 support
- S7-PLCSIM Advanced (for simulation)
- Python 3.10+
- An MQTT broker (for example, a local Mosquitto instance)

Install Python dependencies:

```bash
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r Python_Bridge/requirements.txt
```

Copy `Python_Bridge/config.example.json` to `Python_Bridge/config.json`, then set the simulator address and MQTT settings for your own test environment. Start the bridge with:

```bash
python Python_Bridge/sews_s7_1500_mqtt_bridge.py --config Python_Bridge/config.json
```

Import the SCL files into a **simulation project** in TIA Portal, configure DB100 for non-optimized access, and call `FC_SEWS_NodeRED_DB100` from OB1. Confirm the PLC configuration and network access before running the bridge.

## Safety and data note

This repository is a portfolio/demo selection, not a production PLC program. Use it only in an isolated simulation or test setup. Do not connect it to live equipment. The repository does not include company reports, operational datasets, credentials, or production configuration.
