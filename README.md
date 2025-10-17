# PyVISA CLI for IVI

This is a straightforward CLI written in Python to provide remote access for reading your instrumentation devices. Each command has a docstring that talks about what the command does, what arguments it needs/takes, and an example. If you need specifics on what your instrumentation device supports, please refer to its user manual. This code provides mostly universal commands for each instrumentation device type, i.e. multimeter, oscilloscopes, etc.

### Devices Supported:

- Digital Multimeter (DMM)
- Oscilloscope
- Auto Function Generator (AFG)
- Power Supply Unit (PSU)
- Spectrum Analyzer (SA)
- Vector Network Analyzer (VNA)
- Load Setter/Electronic Load (ELOAD)

### Required Modules:

- pip install pyvisa
- pip install pyvisa_py

### Planned Features:

- More configuration and measurements for every type of instrumentation devices
- More instrumentation devices
